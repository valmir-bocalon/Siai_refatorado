unit RelRecebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, XEdit, Mask, XDate, ExtCtrls, XBanner,
  XLabel3D,   dxButton, DB, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, JvExControls, JvSpecialProgress, Grids, DBGrids, DBClient,
  DBCtrls, TFlatCheckBoxUnit;

type
  TFrmRelRecebimento = class(TForm)
    RGOrdem: TRadioGroup;
    GroupBox1: TGroupBox;
    XDEEntradaInicio: TXDateEdit;
    XDEEntradaFinal: TXDateEdit;
    Label1: TLabel;
    Label2: TLabel;
    XLabel3D1: TXLabel3D;
    RGFormato: TRadioGroup;
    DXBFechar: TdxButton;
    DXBImprimir: TdxButton;
    DS_TempReceber: TDataSource;
    ZQTempReceber: TZQuery;
    GroupBox5: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    XDEVencimentoInicio: TXDateEdit;
    XDEVencimentoFinal: TXDateEdit;
    JSP1: TJvSpecialProgress;
    CBZebrar: TCheckBox;
    GroupBox6: TGroupBox;
    GroupBox7: TGroupBox;
    Label7: TLabel;
    Label10: TLabel;
    XDEBaixasinicio: TXDateEdit;
    XDEBaixasfinal: TXDateEdit;
    CBSaldo: TComboBox;
    DS_Baixa: TDataSource;
    ZQBaixa: TZQuery;
    GroupBox8: TGroupBox;
    CBVerbaixa: TComboBox;
    DS_ResumoReceb: TDataSource;
    ZQResumoReceb: TZQuery;
    DS_SomaBaixa: TDataSource;
    ZQSomaBaixa: TZQuery;

























    DS_AgruTipo: TDataSource;
    ZQAgruTipo: TZQuery;


















    RGContas: TRadioGroup;
    CBObsLanc: TCheckBox;
    CBOBSBaix: TCheckBox;
    DS_Temp2Cl: TDataSource;
    ZQTemp2Cl: TZQuery;
    DS_Cli: TDataSource;
    CDSCli: TClientDataSet;
    CDSCliCCli: TIntegerField;
    CDSCliRcli: TStringField;
    DS_Adversatemp: TDataSource;
    ZQAdvsersatemp: TZQuery;
    DS_Adversatemp2: TDataSource;
    CDSAdversaTemp2: TClientDataSet;
    CDSAdversaTemp2CodAdversa: TIntegerField;
    CDSAdversaTemp2nomeadversa: TStringField;














    DS_Plano: TDataSource;
    CDSPlano: TClientDataSet;
    CDSPlanomascara: TStringField;
    CDSPlanodescricao: TStringField;
    CDSPlanoCrDe: TStringField;
    CDSPlanosoma: TFloatField;
    CDSPlanoclassificacao: TStringField;
    CDSPlanocodigo: TIntegerField;
    DS_Detal: TDataSource;
    CDSDetal: TClientDataSet;
    CDSDetalmascara: TStringField;
    CDSDetaldata: TDateField;
    CDSDetaldetalhe: TMemoField;
    CDSDetalvalor: TFloatField;
    GBCliente: TGroupBox;
    XBanner2: TXBanner;
    ECliente: TEdit;
    DBGrid3: TDBGrid;
    GBAdversa: TGroupBox;
    XBanner10: TXBanner;
    Eadversa: TEdit;
    DBGrid4: TDBGrid;
    DS_Rateio: TDataSource;
    ZQRateio: TZQuery;
    XBanner13: TXBanner;
    XBanner1: TXBanner;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    XBanner5: TXBanner;
    XBanner6: TXBanner;
    XBanner14: TXBanner;
    XBanner15: TXBanner;






























































































































































































































































    CBDesconsidera: TCheckBox;








































    GroupBox10: TGroupBox;
    XBanner16: TXBanner;
    CBEntrada: TFlatCheckBox;
    CBParcela: TFlatCheckBox;
    GroupBox11: TGroupBox;
    XBanner17: TXBanner;
    CLBEmpree: TCheckListBox;
    GroupBox4: TGroupBox;
    XBanner7: TXBanner;
    CLBDoc: TCheckListBox;
    DS_Empree: TDataSource;
    ZQEmpree: TZQuery;



















    DS_Venda: TDataSource;
    ZQVenda: TZQuery;

























    GroupBox2: TGroupBox;
    XBanner8: TXBanner;
    XBanner9: TXBanner;
    EQuadra: TEdit;
    DBGrid1: TDBGrid;
    Elote: TEdit;
    Label3: TLabel;
    DS_QuadraLote: TDataSource;
    CDSQuadraLote: TClientDataSet;
    CDSQuadraLoteQuadra: TStringField;
    CDSQuadraLotelote: TStringField;
    CDSQuadraLoteidlote: TIntegerField;
    CDSQuadraLoteidloteamen: TIntegerField;
    DS_Ven1: TDataSource;
    ZQVen1: TZQuery;

























    DS_ROReceb: TDataSource;
    ZROQReceb: TZReadOnlyQuery;














































    DS_Tipodoc: TDataSource;
    CDSTipoDoc: TClientDataSet;
    CDSTipoDocTipoDoc: TStringField;
    CDSTipoDocdescDoc: TStringField;
    CDSTipoDocdadosch: TStringField;
    DS_ROReceb1: TDataSource;
    ZROQReceb1: TZReadOnlyQuery;




















































    CDSTipoDocsomatit: TFloatField;
    CDSTipoDocsomapago: TFloatField;
    CDSTipoDocsomajuros: TFloatField;



    DS_Ch: TDataSource;
    ZROQCh: TZReadOnlyQuery;





















































    GroupBox3: TGroupBox;
    XBanner11: TXBanner;
    Label4: TLabel;
    Label8: TLabel;
    XDERefInicio: TXDateEdit;
    XDERefFinal: TXDateEdit;
    DataQReceb1: TDataSource;
    QReceb1: TClientDataSet;

    dxbx: TRadioGroup;















    DataZROQReceb2: TDataSource;
    ZROQCh2: TZReadOnlyQuery;
    DataZROQCh2: TDataSource;
























































    ZROQReceb2: TZQuery;

































    QReceb1auto: TAutoIncField;
    QReceb1TipDoc: TStringField;
    Datadocumento: TDataSource;
    documento: TClientDataSet;
    documentoauto: TAutoIncField;
    documentotipdoc: TStringField;
    QReceb1descdoc: TStringField;
    documentodescdoc: TStringField;
    procedure FormShow(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBImprimirClick(Sender: TObject);
    procedure XDEEntradaInicioExit(Sender: TObject);
    procedure XDEVencimentoInicioExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GBEscClienteExit(Sender: TObject);
    procedure ZQBaixaCalcFields(DataSet: TDataSet);
    procedure EClienteExit(Sender: TObject);
    procedure DBGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EadversaExit(Sender: TObject);
    procedure DBGrid4KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RGContasClick(Sender: TObject);
    procedure RGFormatoClick(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure XDERefInicioExit(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmRelRecebimento: TFrmRelRecebimento;

implementation

uses Tabelas, Funcoes, Principal, RelReceb02, Cad_Recebimento,
  RelRecebGerencial, CobrancaBancaria, Recebi03, Recebi04, uRuntimeFields;

{$R *.dfm}

procedure TFrmRelRecebimento.FormShow(Sender: TObject);
begin
  if RGFormato.ItemIndex = 4 Then
    CBDesconsidera.Visible := True
  else
    CBDesconsidera.Visible := False;
  ZQTemp2Cl.Open;
  CDSQuadraLote.Close;
  CDSQuadraLote.CreateDataSet;
  CDSCli.Close;
  CDSCli.CreateDataSet;
  CDSAdversaTemp2.Close;
  CDSAdversaTemp2.CreateDataSet;
  CLBDoc.Sorted := False;
  DM_tabelas.ZQTipodoc.First;
  CLBEmpree.Clear;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQLoteamento.First;
  while not DM_TAbelas.ZQLoteamento.Eof do begin
    CLBEmpree.Items.Add(DM_Tabelas.ZQLoteamento.FieldByName('apelido').Value+' - '+DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Text);
    DM_Tabelas.ZQLoteamento.Next;
  end;
  CLBDoc.Clear;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    CLBDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    CLBDoc.Checked[CLBDoc.Count-1] := True;
   DM_tabelas.ZQTipodoc.Next;
  end;
  CLBDoc.Sorted := True;
  RGContas.SetFocus;
end;

procedure TFrmRelRecebimento.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRelRecebimento.DXBImprimirClick(Sender: TObject);
Var
  VarEmpree, VarDoc, VarOrdem, vartitulo, Varaleatorio, vardevedor, Vrtexto,
  vg,varmasc, vardocumento, varquadra, varsql,VarLogo : String;
  varcontador, x, y, varvezes : Integer;
const
  vartama : array[1..5] of integer = (1,3,6,10,15);
begin
  vg:='';
  VarEmpree:= '';
  Varlogo:='';
  varcontador := 0;
  for x := 0 to CLBEmpree.Count-1 do begin
    if CLBEmpree.Checked[x] Then Begin
      inc(varcontador);
      for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
        if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
          break;
      VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
    end;
  end;
  if varcontador<>CLBEmpree.Count Then
    VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
  else
    VarEmpree := '';
  varlogo:=VarEmpree;
  if not empty(VarLogo) then
  begin
    if pos(',',varlogo)=0 then
    begin
      ZQEmpree.Open;
      ZQEmpree.Filtered:=false;
      ZQEmpree.Filter:='idloteamento='+varlogo;
      ZQEmpree.Filtered:=true;
    end
    else
    begin
      VarLogo:='';
      ZQEmpree.Close;      
    end;
  end;
  varquadra := '';
  CDSQuadraLote.First;
  while not CDSQuadraLote.Eof do begin
    varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
    CDSQuadraLote.Next;
  end;
  varquadra := copy(varquadra,1,length(varquadra)-1);
  vardocumento := '';
  if CBEntrada.Checked <> CBParcela.Checked Then Begin
    if CBParcela.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-P-%');
    if CBEntrada.Checked Then
      vardocumento := ' and documento like '+quotedstr('%-E-%');
  end;
  FrmRelReceb02.XNsomabaixa.Value:=0;
  FrmRelReceb02.XNNegociado.Value:=0;
  FrmRelReceb02.XNRenova.Value:=0;
  FrmRelReceb02.XNPrincipal.Value:=0;
  Varaleatorio := '';
  if CDSCli.Active Then Begin
    CDSCli.First;
    while not CDSCli.Eof do Begin
      Varaleatorio := Varaleatorio +quotedstr(CDSCliCcli.Text)+',';
      CDSCli.Next;
    end;
    Varaleatorio := copy(Varaleatorio,1,length(Varaleatorio)-1)
  end;
  vardevedor := '';
  if CDSAdversaTemp2.Active Then Begin
    CDSAdversaTemp2.First;
    while not CDSAdversaTemp2.Eof do Begin
      vardevedor := vardevedor +quotedstr(CDSAdversaTemp2CodAdversa.Text)+',';
      CDSAdversaTemp2.Next;
    end;
    vardevedor := copy(vardevedor,1,length(vardevedor)-1)
  end;
  case RGOrdem.ItemIndex of
   0: VarOrdem := 'DT_Vencimento';
   1: VarOrdem := 'DT_Entrada';
   2: VArOrdem := 'nome';
  end;
  JSP1.Visible := True;
  VarDoc := '';
  varcontador := 0;
  JSP1.Maximum := CLBEmpree.Count-1;
  for x := 0 to CLBEmpree.Count-1 do begin
    JSP1.Position := x;
    if CLBEmpree.Checked[x] Then Begin
      inc(varcontador);
      for y := length(CLBEmpree.Items.Strings[x]) downto 0 do
        if copy(CLBEmpree.Items.Strings[x],y,1)=' ' Then
          break;
      VarEmpree := VarEmpree +quotedstr(copy(CLBEmpree.Items.Strings[x],y+1,3))+',';
    end;
  end;
  if varcontador<>CLBEmpree.Count Then
    VarEmpree := copy(VarEmpree,1,length(VarEmpree)-1)
  else
    VarEmpree := '';
  varcontador := 0;
  JSP1.Maximum := CLBDoc.Count-1;
  for x := 0 to CLBDoc.Count-1 do begin
    JSP1.Position := x;
    if CLBDoc.Checked[x] Then Begin
      inc(varcontador);
      VarDoc := VarDoc +quotedstr(CLBDoc.Items.Strings[x])+',';
    end;
  end;
  VarDoc := copy(varDoc,1,length(varDoc)-1);
  if RGFormato.ItemIndex = 5 Then Begin
    CDSTipoDoc.Close;
    CDSTipoDoc.CreateDataSet;
    ZROQReceb.SQL.Clear;
    Vrtexto:='';
    ZROQReceb.SQL.Add('SELECT *, sum(Re.valor) as somatit, sum(vr_rec) as somapago, sum(juros_vr) as somajuros FROM receb_baixa as RB');
    ZROQReceb.SQL.Add(' join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento');
    ZROQReceb.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      ZROQReceb.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      ZROQReceb.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZROQReceb.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZROQReceb.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZROQReceb.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
      Vrtexto := 'Entrada de '+XDEEntradaInicio.DateText+' até '+XDEEntradaFinal.DateText;
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Vencimento de '+XDEVencimentoInicio.DateText+' até '+XDEVencimentoFinal.DateText;
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Baixa de '+XDEBaixasInicio.DateText+' até '+XDEBaixasFinal.DateText;
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb.SQL.Add(' and RB.vencimento between :dt7 and :dt8');
      ZROQReceb.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
      if not empty(vrtexto) Then
        vrtexto := vrtexto+' e ';
      Vrtexto := 'Referência de '+XDERefInicio.DateText+' até '+XDERefFinal.DateText;
    end;
    ZROQReceb.SQL.Add(' group by RB.tipdoc ');
    ZROQReceb.Open;
//    if not empty(vrtexto) Then
//      FRM_Recebi03.RLLabel2.Caption := 'Período de : '+vrtexto;
//    FRM_Recebi03.XNumEdit1.Value := 0;
    while not ZROQReceb.Eof do begin
      CDSTipoDoc.Insert;
      CDSTipoDocTipoDoc.Value := ZROQReceb.FieldByName('TipDoc').Value;
      CDSTipoDocsomatit.Value := ZROQReceb.FieldByName('somatit').Value;
      CDSTipoDocsomapago.Value := ZROQReceb.FieldByName('somapago').Value;
      CDSTipoDocsomajuros.Value := ZROQReceb.FieldByName('somajuros').Value;
      CDSTipoDoc.Post;
//      FRM_Recebi03.XNumEdit1.Value := FRM_Recebi03.XNumEdit1.Value+ZROQRecebsomatit.Value;
//      FRM_Recebi03. :=0;
//      FRM_Recebi03.XNumEdit2.Value := FRM_Recebi03.XNumEdit2.Value+ZROQRecebsomapago.Value;
//      FRM_Recebi03.XNumEdit3.Value := FRM_Recebi03.XNumEdit3.Value+ZROQRecebsomajuros.Value;
      ZROQReceb.Next;
    end;
    ZROQReceb1.SQL.Clear;
    ZROQReceb1.SQL.Add('SELECT * FROM receb_baixa as RB join recbxhist as RH on RH.refer=RB.refbaixa join recebimento as RE on RH.idrecib=RE.idrecebimento');
    ZROQReceb1.SQL.Add(' where RB.TIPDOC in ('+varDoc+')');
    if not empty(vardevedor) then
      ZROQReceb1.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      ZROQReceb1.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if not empty(VarEmpree) then
      ZROQReceb1.SQL.Add(' and idloteamento in ('+VarEmpree+')');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZROQReceb1.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZROQReceb1.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZROQReceb1.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZROQReceb1.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
      ZROQReceb1.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      ZROQReceb1.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      ZROQReceb1.SQL.Add(' and RB.vencimento between :dt7 and :dt8');
      ZROQReceb1.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      ZROQReceb1.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;
    ZROQReceb1.Open;
    ZROQCh.Open;
    if CBSaldo.ItemIndex=1 then
    begin
      if FRM_Recebi03=nil then
         FRM_Recebi03:=TFRM_Recebi03.Create(Application);
      FRM_Recebi03.XNumEdit1.Value := 0;
      FRM_Recebi03.XNumEdit2.Value := 0;
      FRM_Recebi03.XNumEdit3.Value := 0;
      if not empty(vrtexto) Then
        FRM_Recebi03.RLLabel2.Caption := 'Período de : '+vrtexto;
      FRM_Recebi03.XNumEdit1.Value := 0;
      ZROQReceb.First;
      while not ZROQReceb.Eof do begin
        FRM_Recebi03.XNumEdit1.Value := FRM_Recebi03.XNumEdit1.Value+ZROQReceb.FieldByName('somatit').Value;
//      FRM_Recebi03. :=0;
        FRM_Recebi03.XNumEdit2.Value := FRM_Recebi03.XNumEdit2.Value+ZROQReceb.FieldByName('somapago').Value;
        FRM_Recebi03.XNumEdit3.Value := FRM_Recebi03.XNumEdit3.Value+ZROQReceb.FieldByName('somajuros').Value;
        ZROQReceb.Next;
      end;
      if dxbx.ItemIndex=1 then
      begin
        FRM_Recebi03.RLLabel15.Visible:=false;
        FRM_Recebi03.RLDBText9.Visible:=false;
      end
      else
      begin
        FRM_Recebi03.RLLabel15.Visible:=true;
        FRM_Recebi03.RLDBText9.Visible:=true;
      end;
      FRM_Recebi03.RLReport1.Preview;
    end
    else
    begin
      if CBSaldo.ItemIndex=0 then
      begin
        if DM_Tabelas.ZQReceb_Baixa.Active= false then
          DM_Tabelas.ZQReceb_Baixa.Open;
        ZROQReceb2.close;
        ZROQReceb2.SQL.Clear;
        ZROQReceb2.SQL.Add('SELECT *  FROM  recebimento  ');
        ZROQReceb2.SQL.Add(' Where TIPDOC in ('+varDoc+')');
        if not empty(vardevedor) then
          ZROQReceb2.SQL.Add(' and adversa in ('+vardevedor+')');
        if not empty(Varaleatorio) then
          ZROQReceb2.SQL.Add(' and cliente in ('+Varaleatorio+')');
        if not empty(VarEmpree) then
          ZROQReceb2.SQL.Add(' and idloteamento in ('+VarEmpree+')');
        if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
          ZROQReceb2.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
          ZROQReceb2.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
          ZROQReceb2.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
        end;
        if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
          ZROQReceb2.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
          ZROQReceb2.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
          ZROQReceb2.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
        end;
{        if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
          ZROQReceb2.SQL.Add(' and RB.DT_Rec between :dt5 and :dt6');
          ZROQReceb2.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
          ZROQReceb2.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
        end;}
        if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
          ZROQReceb2.SQL.Add(' and DT_vencimento between :dt7 and :dt8');
          ZROQReceb2.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
          ZROQReceb2.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
        end;
        ZROQReceb2.Open;
        ZROQReceb2.First;
        if ZROQReceb2.recordcount>0 then
        begin
          QReceb1.close;
          QReceb1.CreateDataSet;
          while not ZROQReceb2.Eof do
          begin
            QReceb1.Insert;
            QReceb1TipDoc.Value:=ZROQReceb2.FieldByName('TipDoc').Value;
            QReceb1.post;
            ZROQReceb2.Next;
          end;
        end
        else
        begin
          showmessage('Nada Encontrado para o Relatório');
          QReceb1.close;
          exit;
        end;
        if QReceb1.RecordCount>0 then
        begin
          QReceb1.IndexName:='tipdocidx';
          QReceb1.First;
          documento.Close;
          documento.CreateDataSet;
          documento.Insert;
          documentotipdoc.Value:=QReceb1TipDoc.value;
          documento.post;
          vg:=QReceb1TipDoc.value;
          while not QReceb1.eof do
          begin
            if vg=QReceb1TipDoc.value then
            begin
              QReceb1.Next;
              vg:=QReceb1TipDoc.value;
            end
            else
            begin
             documento.Insert;
             documentotipdoc.Value:=vg;
             documento.post;
             vg:=QReceb1TipDoc.value;
            end;
          end;
        end;
        QReceb1.Close;
        ZROQReceb2.MasterSource:=Datadocumento;
        ZROQReceb2.MasterFields:='tipdoc';
        ZROQReceb2.LinkedFields:='TipDoc';
        ZROQReceb2.first;
      end
      else if CBSaldo.ItemIndex=2 then
      begin
        //aqui
      end;

      if FRM_Recebi04=nil then
         FRM_Recebi04:=TFRM_Recebi04.Create(Application);
      FRM_Recebi04.XNumEdit1.Value := 0;
      FRM_Recebi04.XNumEdit2.Value := 0;
      FRM_Recebi04.XNumEdit3.Value := 0;
      if not empty(vrtexto) Then
        FRM_Recebi04.RLLabel2.Caption := 'Período de : '+vrtexto;
      FRM_Recebi04.XNumEdit1.Value := 0;
      if dxbx.ItemIndex=1 then
      begin
        FRM_Recebi04.RLLabel15.Visible:=false;
        FRM_Recebi04.RLDBText9.Visible:=false;
      end
      else
      begin
        FRM_Recebi04.RLLabel15.Visible:=true;
        FRM_Recebi04.RLDBText9.Visible:=true;
      end;
      FRM_Recebi04.RLReport1.Preview;
//      DM_Tabelas.ZQReceb_Baixa.close;
    end;
    ZQEmpree.Close;
    documento.Close;
    exit;
  end;
  if RGFormato.ItemIndex = 4 Then Begin
    DM_Tabelas.CDSEmpreendimento.Close;
    DM_Tabelas.CDSEmpreendimento.CreateDataSet;
    DM_Tabelas.ZQLoteamento.First;
    While not DM_Tabelas.ZQLoteamento.Eof do Begin
      while not DM_Tabelas.ZQincorp_loteame.Eof do begin
        if (pos(DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Text,Varaleatorio)>0) or empty(Varaleatorio) Then Begin
          DM_Tabelas.CDSEmpreendimento.Insert;
          DM_TAbelas.CDSEmpreendimentoidordem.Value := DM_Tabelas.CDSEmpreendimento.RecordCount;
          DM_Tabelas.CDSEmpreendimentoidloteamento.Value := DM_Tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
          DM_Tabelas.CDSEmpreendimentocodcontabancaria.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('codcontabancaria').Value;
          DM_Tabelas.CDSEmpreendimentoidparticipante.Value := DM_Tabelas.ZQincorp_loteame.FieldByName('incorporador_idincorporador').Value;
          DM_Tabelas.CDSEmpreendimentomarca.Value := '0';
          DM_Tabelas.CDSEmpreendimento.Post;
        end;
        DM_Tabelas.ZQincorp_loteame.Next;
      end;
      DM_TAbelas.ZQLoteamento.Next;
    end;
    DM_Tabelas.ZQRecebParticp.SQL.Clear;
    DM_Tabelas.ZQRecebParticp.SQL.Add('select * from recebimento join participante on adversa=idpaticipante');
    DM_Tabelas.ZQRecebParticp.SQL.Add(' where recpag = '+quotedstr('R'));
    if CBEntrada.Checked <> CBParcela.Checked Then Begin
      if CBParcela.Checked Then
        DM_Tabelas.ZQRecebParticp.SQL.Add(' and documento like '+quotedstr('%-P-%'));
    end;
    if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and adversa in ('+vardevedor+')');
    if not empty(Varaleatorio) then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and cliente in ('+Varaleatorio+')');
    if CBSaldo.ItemIndex = 1 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and (RE.saldo <re.valor or RE.saldo is null)');
    if CBSaldo.ItemIndex = 2 Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and RE.saldo > 0');
    if not empty(vardocumento) Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(vardocumento);
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
    end;
    if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and BX.vencimento between :dt7 and :dt8');
      DM_Tabelas.ZQRecebParticp.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
      DM_Tabelas.ZQRecebParticp.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
    end;
    if CBDesconsidera.Checked Then
      DM_Tabelas.ZQRecebParticp.SQL.Add(' and not exists (select * from remessa_receb where idrec=idrecebimento)');
    DM_Tabelas.ZQRecebParticp.Open;
    JSP1.Maximum := DM_Tabelas.ZQRecebParticp.RecordCount;
    DM_Tabelas.CDS_MarcaTit.Close;
    DM_TAbelas.CDS_MarcaTit.CreateDataSet;
    while not DM_Tabelas.ZQRecebParticp.Eof do begin
      JSP1.Position := DM_Tabelas.ZQRecebParticp.RecNo;
      DM_Tabelas.CDS_MarcaTit.Insert;
      DM_Tabelas.CDS_MarcaTitIdreceb.Value := DM_Tabelas.ZQRecebParticp.FieldByName('idrecebimento').Value;
      DM_TAbelas.CDS_MarcaTitidparti.Value := DM_Tabelas.ZQRecebParticp.FieldByName('adversa').Value;
      DM_TAbelas.CDS_MarcaTitcliente.Value := DM_Tabelas.ZQRecebParticp.FieldByName('cliente').Value;
      DM_Tabelas.CDS_MarcaTitidloteam.Value := DM_Tabelas.ZQRecebParticp.FieldByName('idloteamento').Value;
      DM_Tabelas.CDS_MarcaTitmarca.Value := '0';
      DM_Tabelas.CDS_MarcaTit.Post;
      DM_Tabelas.ZQRecebParticp.Next;
    end;
    JSP1.Visible := False;
    FrmCobrancaBancaria.Top := FrmRelRecebimento.Top + 40;
    FrmCobrancaBancaria.Left :=FrmRelRecebimento.Left+30;
    FrmCobrancaBancaria.showmodal;
    Exit;
  end;
  vartitulo := 'Ordem de '+RGOrdem.Items.Strings[RGOrdem.ItemIndex];
  if not empty(XDEVencimentoInicio.Text) Then
    vartitulo := vartitulo + ', com vencimento de '+XDEVencimentoInicio.Text+' até '+XDEVencimentoFinal.Text;
  if not empty(XDEEntradaInicio.Text) Then
    vartitulo := vartitulo + ', com Emissão de '+XDEEntradaInicio.Text+' até '+XDEEntradaFinal.Text;
  if not empty(XDEBaixasInicio.Text) Then
    vartitulo := vartitulo + ', com Baixas de '+XDEBaixasInicio.Text+' até '+XDEBaixasFinal.Text;
  vartitulo := vartitulo + ', com Quitados='+CBSaldo.Text+', com Doc.='+VarDoc;
  ZQTempReceber.SQL.Clear;
  ZQTempReceber.SQL.Add('select * from (recebimento as RE join participante as CL on idpaticipante=cliente)');
  if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then
    ZQTempReceber.SQL.Add('join recbxhist as H  ON RE.idrecebimento=H.idrecib join receb_baixa as BX ON BX.refbaixa=H.refer');
  ZQTempReceber.SQL.Add(' where 1 = 1 ');
    if not empty(vardocumento) Then
      ZQTempReceber.SQL.Add(vardocumento);
  if RGFormato.ItemIndex <> 3 Then Begin
    if RGContas.ItemIndex = 0 Then
      ZQTempReceber.SQL.Add(' and recpag ='+quotedstr('R'))
    else
      ZQTempReceber.SQL.Add(' and recpag ='+quotedstr('P'));
  end;
  if not empty(Varaleatorio) then
    ZQTempReceber.SQL.Add(' and idpaticipante in ('+Varaleatorio+')');

  if not empty(VarEmpree) then
     ZQTempReceber.SQL.Add(' and idloteamento in ('+VarEmpree+')');

  if (not empty(vardevedor)) and (RGContas.ItemIndex=0) then
    ZQTempReceber.SQL.Add(' and adversa in ('+vardevedor+')');
  if not empty(varDoc) Then
    ZQTempReceber.SQL.Add(' and TIPDOC in ('+varDoc+')');
  if CBSaldo.ItemIndex = 1 Then
     ZQTempReceber.SQL.Add(' and (RE.saldo <re.valor or RE.saldo is null)');
  if CBSaldo.ItemIndex = 2 Then
     ZQTempReceber.SQL.Add(' and RE.saldo > 0');
  if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
    ZQTempReceber.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
    ZQTempReceber.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
  end;
  if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
    ZQTempReceber.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
    ZQTempReceber.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
  end;
  if  not empty(copy(XDEBaixasInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and BX.DT_Rec between :dt5 and :dt6');
    ZQTempReceber.ParamByName('dt5').AsDate:=strtodate(XDEBaixasInicio.DateText);
    ZQTempReceber.ParamByName('dt6').AsDate:=strtodate(XDEBaixasFinal.DateText);
  end;
  if  not empty(copy(XDERefInicio.Text,1,2)) Then Begin
    ZQTempReceber.SQL.Add(' and BX.DT_Rec between :dt7 and :dt8');
    ZQTempReceber.ParamByName('dt7').AsDate:=strtodate(XDERefInicio.DateText);
    ZQTempReceber.ParamByName('dt8').AsDate:=strtodate(XDERefFinal.DateText);
  end;
  ZQTempReceber.SQL.Add(' group by re.idrecebimento');
  ZQTempReceber.SQL.Add(' ORDER BY '+VarOrdem);
  ZQTempReceber.Open;
  //2
  if RGFormato.ItemIndex = 2 Then Begin
    ZQResumoReceb.SQL.Clear;
    ZQResumoReceb.SQL.Add('select * from recebimento as RE join participante as CL on idpaticipante=cliente ');
    if RGContas.ItemIndex = 0 Then
      ZQResumoReceb.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQResumoReceb.SQL.Add(' where recpag ='+quotedstr('P'));
    if not empty(vardocumento) Then
      ZQResumoReceb.SQL.Add(vardocumento);
    if not empty(varDoc) Then
      ZQResumoReceb.SQL.Add(' and TIPDOC in ('+varDoc+')');
    if CBSaldo.ItemIndex = 1 Then
       ZQResumoReceb.SQL.Add(' and RE.saldo > 0');
    if CBSaldo.ItemIndex = 2 Then
       ZQResumoReceb.SQL.Add(' and (RE.saldo = 0 or RE.saldo is null)');
    if  not empty(copy(XDEEntradaInicio.Text,1,2)) Then Begin
      ZQResumoReceb.SQL.Add(' and DT_Entrada between :dt1 and :dt2');
      ZQResumoReceb.ParamByName('dt1').AsDate:=strtodate(XDEEntradaInicio.DateText);
      ZQResumoReceb.ParamByName('dt2').AsDate:=strtodate(XDEEntradaFinal.DateText);
    end;
    if  not empty(copy(XDEVencimentoInicio.Text,1,2)) Then Begin
      ZQResumoReceb.SQL.Add(' and DT_Vencimento between :dt3 and :dt4');
      ZQResumoReceb.ParamByName('dt3').AsDate:=strtodate(XDEVencimentoInicio.DateText);
      ZQResumoReceb.ParamByName('dt4').AsDate:=strtodate(XDEVencimentoFinal.DateText);
    end;
    ZQResumoReceb.SQL.Add(' GROUP BY RE.tipdoc ORDER BY '+VarOrdem);
    ZQResumoReceb.Open;
    ZQSomaBaixa.SQL.Clear;
    ZQSomaBaixa.SQL.Add('select *, sum(vr_rec) as somabaixa FROM (recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib)');
    ZQSomaBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
    if RGContas.ItemIndex = 0 Then
      ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQSomaBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQSomaBaixa.SQL.Add(' group by b.refbaixa');
    ZQSomaBaixa.Open;
    ZQBaixa.SQL.Clear;
    ZQBaixa.SQL.Add(' SELECT * FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib');
    ZQBaixa.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
    ZQBaixa.SQL.Add(' join tipodocumento ON tipodoc= b.tipdoc');
    if RGContas.ItemIndex = 0 Then
      ZQBaixa.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQBaixa.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQBaixa.Open;
    ZQAgruTipo.SQL.Clear;
    ZQAgruTipo.SQL.Add('Select *, sum(saldo) as somasaldo from recebimento');
    if RGContas.ItemIndex = 0 Then
      ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('R'))
    else
      ZQAgruTipo.SQL.Add(' where recpag ='+quotedstr('P'));
    ZQAgruTipo.SQL.Add(' group by TipDoc');
    ZQAgruTipo.Open;
    FrmRelReceb02.RLDBResult1.DataField := 'saldo';
    FrmRelReceb02.RLDBText9.DataField := 'saldo';
    FrmRelReceb02.RLLabel1.Caption := 'Contas a Receber';
    if RGContas.ItemIndex = 1 Then
    FrmRelReceb02.RLLabel1.Caption := 'Contas a Pagar';
    if FrmRelReceb02=nil then
      FrmRelReceb02:=TFrmRelReceb02.Create(Application);
    FrmRelReceb02.RLReport1.Preview;
    ZQEmpree.Close;
  end
  else if RGFormato.ItemIndex = 3 Then Begin
    CDSPlano.Close;
    CDSPlano.CreateDataSet;
    CDSPlano.IndexFieldNames := 'mascara';
    CDSDetal.Close;
    CDSDetal.CreateDataSet;
    CDSDetal.IndexFieldNames := 'data';
    CDSDetal.Filtered := True;
    JSP1.Visible := True;
    JSP1.Maximum := ZQTempReceber.RecordCount;
    while not ZQTempReceber.Eof do begin
      JSP1.Position := ZQTempReceber.RecNo;
      for varvezes:=1 to 5 do begin
        varmasc := copy(ZQTempReceber.FieldByName('planomascara').Value,1,vartama[varvezes]);
        if length(varmasc)<>vartama[varvezes] Then
          break;
        if CDSPlano.Locate('mascara',varmasc,[]) Then
          CDSPlano.Edit
        else
          CDSPlano.Insert;
        CDSPlanomascara.Value := varmasc;
        CDSPlanosoma.Value := CDSPlanosoma.Value + ZQTempReceber.FieldByName('Valor').Value;
        CDSPlano.Post;
      end;
      CDSDetal.Insert;
      CDSDetalmascara.Value := ZQTempReceber.FieldByName('planomascara').Value;
      CDSDetalvalor.Value := ZQTempReceber.FieldByName('Valor').Value;
      CDSDetaldata.Value := ZQTempReceber.FieldByName('Dt_Vencimento').Value;
      CDSDetaldetalhe.Value := 'Lanç: '+ZQTempReceber.FieldByName('idrecebimento').Text+'  Participante: '+ZQTempReceber.FieldByName('nome_parte').Value+'      Adversa: '+ZQTempReceber.FieldByName('adversanome').Value+chr(13);
      CDSDetaldetalhe.Value := CDSDetaldetalhe.Value + ZQTempReceber.FieldByName('Observ').Value;
      ZQTempReceber.Next;
    end;
    if CBObsLanc.Checked Then
      FrmRelRecebGerencial.RLBand4.Visible := True
    else
      FrmRelRecebGerencial.RLBand4.Visible := False;
    if FrmRelRecebGerencial=nil then
      FrmRelRecebGerencial:=TFrmRelRecebGerencial.Create(Application);
    FrmRelRecebGerencial.RLReport1.Preview;
    ZQEmpree.Close;    
  end;
  JSP1.Visible := False;
end;

procedure TFrmRelRecebimento.XDEEntradaInicioExit(Sender: TObject);
begin
  if XDEBaixasfinal.DateValue<XDEBaixasinicio.DateValue Then
     XDEBaixasfinal.DateValue := XDEBaixasinicio.DateValue;
end;

procedure TFrmRelRecebimento.XDEVencimentoInicioExit(Sender: TObject);
begin
  if XDEVencimentoFinal.DateValue<XDEVencimentoInicio.DateValue Then
     XDEVencimentoFinal.DateValue := XDEVencimentoInicio.DateValue;
end;

procedure TFrmRelRecebimento.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if not (ActiveControl is TDBGrid) then begin
      Key := #0;
      Perform(WM_NEXTDLGCTL, 0, 0);
    end;
end;

procedure TFrmRelRecebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQBaixa.Close;
  ZQTempReceber.Close;
end;

procedure TFrmRelRecebimento.GBEscClienteExit(Sender: TObject);
begin
  FrmRelRecebimento.KeyPreview := True
end;

procedure TFrmRelRecebimento.ZQBaixaCalcFields(DataSet: TDataSet);
begin
  if ZQBaixa.FieldByName('percent_usado').Value >0 Then
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value*ZQBaixa.FieldByName('percent_usado').Value/100
  else
    ZQBaixa.FieldByName('vrusado').Value := ZQBaixa.FieldByName('Vr_rec').Value;
end;

procedure TFrmRelRecebimento.EClienteExit(Sender: TObject);
begin
  if not empty(ECliente.Text) Then Begin
    if AchaCLiente(200, 10,ECliente.Text) Then Begin
      CDSCli.Insert;
      CDSCliCCli.Value := DM_TAbelas.ZQAchaParticip.FieldByName('idpaticipante').Value;
      CDSCli.Post;
    end;
    ECliente.Text := '';
    ECliente.SetFocus;
  end;
end;

procedure TFrmRelRecebimento.DBGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSCli.RecordCount>0) Then Begin
    CDSCli.Delete;
  end;
end;

procedure TFrmRelRecebimento.EadversaExit(Sender: TObject);
begin
  IF NOT EMPTY(Eadversa.Text) Then Begin
    if AchaCliente(400, 300, Eadversa.Text) then Begin
      CDSAdversaTemp2.Insert;
      CDSAdversaTemp2CodAdversa.Value := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Value;
      CDSAdversaTemp2.Post;
    end;
    Eadversa.Text := '';
    Eadversa.SetFocus;
  end;
end;

procedure TFrmRelRecebimento.DBGrid4KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;
end;

procedure TFrmRelRecebimento.RGContasClick(Sender: TObject);
begin
  if RGContas.ItemIndex = 0 then
    GBAdversa.Enabled := True
  else
    GBAdversa.Enabled := False;
end;

procedure TFrmRelRecebimento.RGFormatoClick(Sender: TObject);
begin
if RGFormato.ItemIndex = 5 Then
   dxbx.Visible:=true
else
   dxbx.Visible:=false;
if RGFormato.ItemIndex = 4 Then
  CBDesconsidera.Visible := True
else
  CBDesconsidera.Visible := False;
end;

procedure TFrmRelRecebimento.EloteExit(Sender: TObject);
begin
  if not empty(Elote.Text+EQuadra.Text) Then Begin
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Add('Select * from venda join imovel on idimovel=imovel');
    ZQVen1.SQL.Add(' where quadra='+quotedstr(Equadra.Text)+' and lote='+quotedstr(Elote.Text));
    ZQVen1.Open;
    if ZQVen1.RecordCount>0 Then Begin
      CDSQuadraLote.Insert;
      CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').Value;
      CDSQuadraLoteQuadra.Value := Equadra.Text+'-'+Elote.Text;
      CDSQuadraLote.Post;
      if ZQVen1.FieldByName('datavenda').Value<XDEEntradaInicio.DateValue Then
        XDEEntradaInicio.DateValue := ZQVen1.FieldByName('datavenda').Value;
      if ZQVen1.FieldByName('datavenda').Value>XDEEntradaFinal.DateValue Then
        XDEEntradaFinal.DateValue := ZQVen1.FieldByName('datavenda').Value;
    end
    else
      Showmessage('Esta Quadra e lote não foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end
  else
    if CDSQuadraLote.RecordCount>0 Then
      DXBImprimir.SetFocus;
end;

procedure TFrmRelRecebimento.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;
end;

procedure TFrmRelRecebimento.XDERefInicioExit(Sender: TObject);
begin
  if XDERefFinal.DateValue<XDERefInicio.DateValue Then
     XDERefFinal.DateValue := XDERefInicio.DateValue;
end;


procedure TFrmRelRecebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'nomeloteamento', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQBaixa', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaTipdoc_1', 'Tipdoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadescricao', 'descricao', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'codcontabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'descricao', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixatipodoc', 'tipodoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_receb', 'vend_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavend_caixa', 'vend_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_caixa', 'receb_caixa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixareceb_receb', 'receb_receb', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadados_chequ', 'dados_chequ', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaso_avista', 'so_avista', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancabanco', 'lancabanco', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasomapaga', 'somapaga', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixalancinclus', 'lancinclus', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavrusado', 'vrusado', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidtipodocumento', 'idtipodocumento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQBaixa', 'ZQBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQResumoReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebtipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoReceblocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebnumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQResumoReceb', 'ZQResumoRecebSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQSomaBaixa', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasomabaixa', 'somabaixa', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQSomaBaixa', 'ZQSomaBaixasubstituicao_1', 'substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAgruTipo', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidRecebimento', 'idRecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiposomasaldo', 'somasaldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiporefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTiponumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAgruTipo', 'ZQAgruTipoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQTemp2Cl', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cldoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cltipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Claniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Cllocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2ClcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQTemp2Cl', 'ZQTemp2Clcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQRateio', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiotipodocumento', 'tipodocumento', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateiovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateionumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQRateio', 'ZQRateioSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQEmpree', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreematriculaloteamento', 'matriculaloteamento', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreedataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreelogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreemapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha1', 'testemunha1', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha2', 'testemunha2', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha3', 'testemunha3', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetestemunha4', 'testemunha4', TStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeObservincorp', 'Observincorp', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreebairro', 'bairro', TStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreepasta_mapa', 'pasta_mapa', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreenomeloteamento', 'nomeloteamento', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreetipo', 'tipo', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreecodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQEmpree', 'ZQEmpreeapelido', 'apelido', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVenda', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaquadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendalado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendamedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendavalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendaObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendacornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVenda', 'ZQVendadisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZQVen1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1quadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1lado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomatit', 'somatit', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomapago', 'somapago', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsomajuros', 'somajuros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeadversa', 'nomeadversa', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQReceborigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebnumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebsubstituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb', 'ZROQRecebSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb1', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Docum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1obsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Codcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1vencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1refer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1valor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1descon', 'descon', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1juros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1percent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1data', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Valor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1TipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1RefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1numboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeadversa', 'nomeadversa', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomecliente', 'nomecliente', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1nomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1substituicao', 'substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb1', 'ZROQReceb1Substituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQCh', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChbanco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChagencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChalias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChconta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChmarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChnumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChTipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDocum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChRefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChobsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChvencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChSubstituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh', 'ZROQChsubstituicao_2', 'substituicao_2', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQCh2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2banco', 'banco', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2agencia', 'agencia', TStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2dono', 'dono', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2CPF_CNPJ', 'CPF_CNPJ', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2valor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2emissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2deposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2alias', 'alias', TStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2prorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2conta', 'conta', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2numero', 'numero', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Substituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Valor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2numboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Substituicao_1', 'Substituicao_1', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2idreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Dt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Vr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2TipDoc_1', 'TipDoc_1', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Docum', 'Docum', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Juros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Desc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2RefBaixa_1', 'RefBaixa_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2obsebx', 'obsebx', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2Codcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2dataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2vencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQCh2', 'ZROQCh2substituicao_2', 'substituicao_2', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRelRecebimento, 'ZROQReceb2', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2idrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2documento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2cliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2usuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Dt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Dt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Valor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Observ', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2VrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2ordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2TipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2saldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2marcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2RefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2refvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2contabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2empresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2custodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2origem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2adversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2recpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2numordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2idloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2venda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2quadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2numboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Substituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2nomeempree', 'nomeempree', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Dt_rec', 'Dt_rec', TDateField, fkLookup, 0, 0, False, '', '', '', '', 0, 'RefBaixa', 'DM_Tabelas.ZQReceb_Baixa', 'RefBaixa', 'Dt_rec', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Vencimento', 'Vencimento', TDateField, fkLookup, 0, 0, False, '', '', '', '', 0, 'RefBaixa', 'DM_Tabelas.ZQReceb_Baixa', 'RefBaixa', 'vencimento', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Juros_Vr', 'Juros_Vr', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'RefBaixa', 'DM_Tabelas.ZQReceb_Baixa', 'RefBaixa', 'Juros_Vr', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2Vr_rec', 'Vr_rec', TFloatField, fkLookup, 0, 0, False, '', '###,###,##0.00', '', '', 0, 'RefBaixa', 'DM_Tabelas.ZQReceb_Baixa', 'RefBaixa', 'Vr_rec', True);
  RegisterRuntimeField(TFrmRelRecebimento, 'ZROQReceb2', 'ZROQReceb2nomeadverda', 'nomeadversa', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);

end.
