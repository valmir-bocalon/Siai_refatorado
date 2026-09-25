unit ReajusteDeParcelas2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, XNum, XLabel3D, Mask, XDate, StdCtrls, Grids, DBGrids,
    dxButton, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBCtrls, TFlatGaugeUnit,DbiProcs,Shellapi, ImgList, ExtCtrls,
  ComCtrls;

type
  TFrm_ReajusteDeParcelas2 = class(TForm)
    XBanner4: TXBanner;
    GBAdversa: TGroupBox;
    XBanner10: TXBanner;
    XBanner6: TXBanner;
    DBGrid4: TDBGrid;
    GroupBox5: TGroupBox;
    XBanner3: TXBanner;
    XLabel3D1: TXLabel3D;
    DXBFechar: TdxButton;
    dxButton1: TdxButton;
    DS_quadraLote: TDataSource;
    DS_Adversatemp: TDataSource;
    ZQAdvsersatemp: TZQuery;

































    CDSAdversaTemp2: TClientDataSet;
    CDSAdversaTemp2CodAdversa: TIntegerField;
    CDSAdversaTemp2nomeadversa: TStringField;
    DataCDSAdversaTEmp2: TDataSource;
    CDSQuadraLote: TClientDataSet;
    CDSQuadraLoteQuadra: TStringField;
    CDSQuadraLotelote: TStringField;
    CDSQuadraLoteidlote: TIntegerField;
    CDSQuadraLoteidloteamen: TIntegerField;
    DS_Ven1: TDataSource;
    ZQVen1: TZQuery;

























    DS_TempReceber: TDataSource;
    ZQTempReceber: TZQuery;
































































    GroupBox2: TGroupBox;
    XBanner8: TXBanner;
    XBanner9: TXBanner;
    Label3: TLabel;
    EQuadra: TEdit;
    DBGrid1: TDBGrid;
    Elote: TEdit;
    percent: TXNumEdit;
    Eadversa: TEdit;
    DBGReceb: TDBGrid;
    DXBPrimeiro: TdxButton;
    DXBAnterior: TdxButton;
    DXBProximo: TdxButton;
    DXBUltimo: TdxButton;
    DXBPesquisar: TdxButton;
    dxButton2: TdxButton;
    DBEIDReceb: TDBEdit;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    xmes: TComboBox;
    xano: TComboBox;
    lmes: TLabel;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;


































    DBGrid2: TDBGrid;
    ZQRecebtmp: TClientDataSet;
    DataZQRecebtmp: TDataSource;
    ZQRecebtmpsaldocalc: TFloatField;
    ZQRecebtmpidrecebimento: TLargeintField;
    ZQRecebtmpdocumento: TStringField;
    ZQRecebtmpcliente: TIntegerField;
    ZQRecebtmpusuario: TIntegerField;
    ZQRecebtmpDt_Entrada: TDateField;
    ZQRecebtmpDt_Vencimento: TDateField;
    ZQRecebtmpValor: TFloatField;
    ZQRecebtmpObserv: TMemoField;
    ZQRecebtmpVrDoc: TFloatField;
    ZQRecebtmpordem: TStringField;
    ZQRecebtmpTipDoc: TStringField;
    ZQRecebtmpsaldo: TFloatField;
    ZQRecebtmpmarcar: TStringField;
    ZQRecebtmpRefBaixa: TIntegerField;
    ZQRecebtmprefvinda: TIntegerField;
    ZQRecebtmpcontabil: TIntegerField;
    ZQRecebtmpempresa: TIntegerField;
    ZQRecebtmpcustodaparcela: TFloatField;
    ZQRecebtmporigem: TStringField;
    ZQRecebtmpadversa: TIntegerField;
    ZQRecebtmprecpag: TStringField;
    ZQRecebtmpnumordem: TIntegerField;
    ZQRecebtmpnomecli: TStringField;
    ZQRecebtmpadversanome: TStringField;
    ZQRecebtmpidloteamento: TIntegerField;
    ZQRecebtmpnome_loteamento: TStringField;
    ZQRecebtmpvenda_idvenda: TIntegerField;
    ZQRecebtmpquadralote: TStringField;
    ZQRecebtmpnumboleto: TStringField;
    ZQRecebtmpSubstituicao: TStringField;
    ZQRecebtmpsq: TLargeintField;
    ZQRecebtmpnomeadversa: TStringField;
    ZQRecebtmpsomar: TStringField;
    ZQRecebtmpsaldo_corr: TFloatField;
    DBid: TDBEdit;
    ListBox1: TListBox;


    dxButton3: TdxButton;
    DS_Entrada: TDataSource;
    ZQEntrada: TZQuery;































    DS_Parcela: TDataSource;
    ZQParcela: TZQuery;





























    ZQVenda: TZQuery;





    DS_Venda: TDataSource;
    ZQEmpree: TZQuery;























    DS_Empree: TDataSource;
    CDSLoteamento: TClientDataSet;
    DataCDSLoteamento: TDataSource;
    CDSLoteamentoidloteamento: TLargeintField;
    CDSLoteamentonomeloteamento: TStringField;
    CDSLoteamentocodcontabancaria: TIntegerField;
    CDSLoteamentoapelido: TStringField;
    GroupBox3: TGroupBox;
    XBanner2: TXBanner;
    XBanner5: TXBanner;
    DBGrid3: TDBGrid;
    Eloteamento: TEdit;
    dxButton14: TdxButton;
    dxButton13: TdxButton;

    ZQRecebtmpproximo_reajuste: TStringField;

    ZQRecebtmpParcelas_fixas: TStringField;
    CBDesconsidera: TCheckBox;
    ImageList1: TImageList;
    dxButton4: TdxButton;
    dxButton5: TdxButton;
    Image1: TImage;
    Image2: TImage;
    DBGrid5: TDBGrid;
    DataZQRecebtmp2: TDataSource;
    ZQRecebtmp2: TClientDataSet;
    ZQRecebtmp2saldocalc: TFloatField;
    ZQRecebtmp2idrecebimento: TLargeintField;
    ZQRecebtmp2documento: TStringField;
    ZQRecebtmp2cliente: TIntegerField;
    ZQRecebtmp2usuario: TIntegerField;
    ZQRecebtmp2Dt_Entrada: TDateField;
    ZQRecebtmp2Dt_Vencimento: TDateField;
    ZQRecebtmp2Valor: TFloatField;
    ZQRecebtmp2Observ: TMemoField;
    ZQRecebtmp2VrDoc: TFloatField;
    ZQRecebtmp2ordem: TStringField;
    ZQRecebtmp2TipDoc: TStringField;
    ZQRecebtmp2saldo: TFloatField;
    ZQRecebtmp2marcar: TStringField;
    ZQRecebtmp2RefBaixa: TIntegerField;
    ZQRecebtmp2refvinda: TIntegerField;
    ZQRecebtmp2contabil: TIntegerField;
    ZQRecebtmp2empresa: TIntegerField;
    ZQRecebtmp2custodaparcela: TFloatField;
    ZQRecebtmp2origem: TStringField;
    ZQRecebtmp2adversa: TIntegerField;
    ZQRecebtmp2recpag: TStringField;
    ZQRecebtmp2numordem: TIntegerField;
    ZQRecebtmp2nomecli: TStringField;
    ZQRecebtmp2adversanome: TStringField;
    ZQRecebtmp2idloteamento: TIntegerField;
    ZQRecebtmp2nome_loteamento: TStringField;
    ZQRecebtmp2venda_idvenda: TIntegerField;
    ZQRecebtmp2quadralote: TStringField;
    ZQRecebtmp2numboleto: TStringField;
    ZQRecebtmp2Substituicao: TStringField;
    ZQRecebtmp2sq: TLargeintField;
    ZQRecebtmp2nomeadversa: TStringField;
    ZQRecebtmp2somar: TStringField;
    ZQRecebtmp2saldo_corr: TFloatField;
    ZQRecebtmp2proximo_reajuste: TStringField;
    ZQRecebtmp2Parcelas_fixas: TStringField;
    barra2: TProgressBar;
    cbconferencia: TCheckBox;

    procedure DXBFecharClick(Sender: TObject);
    procedure EadversaExit(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure botoes_setas;
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid4KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dxButton2Click(Sender: TObject);
    procedure DBGRecebDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure dxButton1Click(Sender: TObject);
    procedure xmesExit(Sender: TObject);
    procedure dxButton13Click(Sender: TObject);
    procedure dxButton14Click(Sender: TObject);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure dxButton3Click(Sender: TObject);
    procedure EloteamentoExit(Sender: TObject);
    procedure GroupBox5Enter(Sender: TObject);
    procedure percentExit(Sender: TObject);
    procedure DBGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dxButton4Click(Sender: TObject);
    procedure dxButton5Click(Sender: TObject);
    procedure DBGrid5DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid5MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure cbconferenciaClick(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_ReajusteDeParcelas2: TFrm_ReajusteDeParcelas2;

implementation

uses tabelas, PesqRecebimento,funcoes, AchaIgpm, PesqRecebimento2,relajuste, uRuntimeFields;

{$R *.dfm}


procedure GetSelectedRecord(DataSet: TDataSet; FieldName: String;
                            BookMarkList: TBookMarkList; Strings : TStrings);
var i,t: Integer;
    bm: TBookMark;
begin
  Strings.Clear;
  With DataSet do
  begin
    DisableControls;
    bm := GetBookMark;
    try
      Frm_ReajusteDeParcelas2.barra2.Position:=0;
      t:=BookMarkList.Count -1;
      Frm_ReajusteDeParcelas2.barra2.Max:=t;
      for i := 0 to BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Frm_ReajusteDeParcelas2.barra2.Position:=i;
        Strings.Add(DataSet.FindField(FieldName).AsString);
        Frm_ReajusteDeParcelas2.ZQRecebimento.Filtered:=false;
        Frm_ReajusteDeParcelas2.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        Frm_ReajusteDeParcelas2.ZQRecebimento.Filtered:=true;
        Frm_ReajusteDeParcelas2.ZQRecebimento.First;
        if Frm_ReajusteDeParcelas2.ZQRecebimento.recordcount>0 then
        begin
  //        Frm_ReajusteDeParcelas2.barra.MaxValue:=Frm_ReajusteDeParcelas2.ZQRecebimento.recordcount;
          while not Frm_ReajusteDeParcelas2.ZQRecebimento.Eof do
          begin
//            Frm_ReajusteDeParcelas2.barra.Progress:=Frm_ReajusteDeParcelas2.ZQRecebimento.recno;
            Frm_ReajusteDeParcelas2.ZQRecebtmp.Insert;
            Frm_ReajusteDeParcelas2.ZQRecebtmpidrecebimento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('idrecebimento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpdocumento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('documento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpcliente.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('cliente').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpusuario.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('usuario').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpDt_Entrada.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Dt_Entrada').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpDt_Vencimento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Dt_Vencimento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpValor.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Valor').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpVrDoc.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('VrDoc').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpordem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('ordem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpTipDoc.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('TipDoc').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpsaldo.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('saldo').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmporigem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('origem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpadversa.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('adversa').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmprecpag.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('recpag').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpnumordem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('numordem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpadversanome.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('adversanome').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpidloteamento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('idloteamento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpvenda_idvenda.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('venda_idvenda').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpquadralote.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('quadralote').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpnomeadversa.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('nomeadversa').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpParcelas_fixas.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Parcelas_fixas').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmpnome_loteamento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('nome_loteamento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp.Post;


            Frm_ReajusteDeParcelas2.ZQRecebtmp2.Insert;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2idrecebimento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('idrecebimento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2documento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('documento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2cliente.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('cliente').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2usuario.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('usuario').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2Dt_Entrada.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Dt_Entrada').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2Dt_Vencimento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Dt_Vencimento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2Valor.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Valor').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2VrDoc.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('VrDoc').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2ordem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('ordem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2TipDoc.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('TipDoc').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2saldo.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('saldo').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2origem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('origem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2adversa.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('adversa').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2recpag.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('recpag').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2numordem.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('numordem').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2adversanome.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('adversanome').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2idloteamento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('idloteamento').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2venda_idvenda.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('venda_idvenda').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2quadralote.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('quadralote').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2nomeadversa.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('nomeadversa').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2Parcelas_fixas.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('Parcelas_fixas').Value;
            Frm_ReajusteDeParcelas2.ZQRecebtmp2nome_loteamento.Value:=Frm_ReajusteDeParcelas2.ZQRecebimento.FieldByName('nome_loteamento').Value;            
            Frm_ReajusteDeParcelas2.ZQRecebtmp2.Post;

            Frm_ReajusteDeParcelas2.ZQRecebimento.Next;
          end;
        end;
        Frm_ReajusteDeParcelas2.ZQRecebimento.Filtered:=false;
        Frm_ReajusteDeParcelas2.barra2.Position:=0;
      end;
    finally
      GotoBookMark(bm);
      FreeBookMark(bm);
      Frm_ReajusteDeParcelas2.barra2.Position:=0;
      EnableControls;
    end;
  end; // With DataSet do
end;


procedure TFrm_ReajusteDeParcelas2.DXBFecharClick(Sender: TObject);
begin
  close;
end;
procedure TFrm_ReajusteDeParcelas2.EadversaExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelas2.EloteExit(Sender: TObject);
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
    end
    else
      Mensagem('Esta Quadra e lote não foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end;
end;

procedure TFrm_ReajusteDeParcelas2.FormShow(Sender: TObject);
var
mmes,ms,xdata:string;
ds:Tdatetime;

begin
  ms:=mesano(date);
  ms:=UltimoDiaDoMes(ms);
  xdata:=ms+copy(datetostr(date),3,10);
  ds:=strtodate(xdata);
  xano.text:=copy(datetostr(date),7,4);
  mmes:=mes(date);
  if mmes='01' then
     xmes.ItemIndex:=0;
  if mmes='02' then
     xmes.ItemIndex:=1;
  if mmes='03' then
     xmes.ItemIndex:=2;
  if mmes='04' then
     xmes.ItemIndex:=3;
  if mmes='05' then
     xmes.ItemIndex:=4;
  if mmes='06' then
     xmes.ItemIndex:=5;
  if mmes='07' then
     xmes.ItemIndex:=6;
  if mmes='08' then
     xmes.ItemIndex:=7;
  if mmes='09' then
     xmes.ItemIndex:=8;
  if mmes='10' then
     xmes.ItemIndex:=9;
  if mmes='11' then
     xmes.ItemIndex:=10;
  if mmes='12' then
     xmes.ItemIndex:=11;
  lmes.Caption:=mmes;


  ZQRecebimento.close;
  ZQRecebimento.SQL.clear;
{  ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
  ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
  ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

{  if CBDesconsidera.Checked=false then
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');}


  if CBDesconsidera.Checked=false then
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento');

  ZQRecebimento.Open;
  ZQRecebimento.First;
  ZQTempReceber.Open;
  ZQVen1.Open;
  ZQAdvsersatemp.Open;
  CDSAdversaTemp2.CreateDataSet;
  CDSQuadraLote.CreateDataSet;
  CDSLoteamento.CreateDataSet;
  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;

  ZQRecebtmp.close;
  ZQRecebtmp.createdataset;
  ZQRecebtmp2.close;
  ZQRecebtmp2.createdataset;


end;
procedure TFrm_ReajusteDeParcelas2.botoes_setas;
Begin
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  if ZQRecebimento.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  if ZQRecebimento.RecNo = ZQRecebimento.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
End;

procedure TFrm_ReajusteDeParcelas2.DXBPrimeiroClick(Sender: TObject);
begin
  ZQRecebimento.First;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas2.DXBAnteriorClick(Sender: TObject);
begin
  ZQRecebimento.Prior;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas2.DXBProximoClick(Sender: TObject);
begin
  ZQRecebimento.Next;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas2.DXBUltimoClick(Sender: TObject);
begin
  ZQRecebimento.Last;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas2.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento2.Top := Frm_ReajusteDeParcelas2.Top+100;
  FrmPesqRecebimento2.Left := Frm_ReajusteDeParcelas2.Left;
  FrmPesqRecebimento2.showmodal;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas2.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrm_ReajusteDeParcelas2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  ZQRecebimento.close;
  ZQTempReceber.close;
  ZQVen1.Close;
  ZQAdvsersatemp.Close;
  CDSAdversaTemp2.Close;
  CDSLoteamento.Close;
  CDSQuadraLote.Close;
  ZQRecebtmp.close;
end;

procedure TFrm_ReajusteDeParcelas2.DBGrid4KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelas2.DBGrid1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelas2.dxButton2Click(Sender: TObject);
var
varquadra,vardevedor,varlotea:string;
mmes,ms,xdata:string;
ds:Tdatetime;
begin
  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Março' then
     xdata:='03';
  if mmes='Abril' then
     xdata:='04';
  if mmes='Maio' then
     xdata:='05';
  if mmes='Junho' then
     xdata:='06';
  if mmes='Julho' then
     xdata:='07';
  if mmes='Agosto' then
     xdata:='08';
  if mmes='Setembro' then
     xdata:='09';
  if mmes='Outubro' then
     xdata:='10';
  if mmes='Novembro' then
     xdata:='11';
  if mmes='Dezembro' then
     xdata:='12';
  lmes.Caption:=xdata;

  ms:=UltimoDiaDoMes(xdata+xano.text);

  varquadra := '';
  if CDSQuadraLote.Active Then
  Begin
     CDSQuadraLote.First;
     if CDSQuadraLote.RecordCount>0 then
     begin
       while not CDSQuadraLote.Eof do begin
         varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
         CDSQuadraLote.Next;
       end;
       varquadra := copy(varquadra,1,length(varquadra)-1);
       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento,ordem,quadralote')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento,ordem,quadralote');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;

  vardevedor := '';
  if CDSAdversaTemp2.Active Then
  Begin
    CDSAdversaTemp2.First;
    if CDSAdversaTemp2.RecordCount>0 then
    begin
      while not CDSAdversaTemp2.Eof do Begin
        vardevedor := vardevedor +quotedstr(CDSAdversaTemp2CodAdversa.Text)+',';
        CDSAdversaTemp2.Next;
      end;
      vardevedor := copy(vardevedor,1,length(vardevedor)-1);
//      ms:=mesano(date);
 //     ms:=UltimoDiaDoMes(ms);
 //     xdata:=ms+copy(datetostr(date),3,10);
//      ds:=strtodate(xdata);
      ZQRecebimento.close;
      ZQRecebimento.SQL.clear;
{      ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and adversa in ('+vardevedor+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
      ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
      ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
      if CBDesconsidera.Checked=false then
         ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento,ordem,adversa')
      else
         ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento,ordem,adversa');
      ZQRecebimento.Open;
      ZQRecebimento.First;
    end;
  end;

  varlotea:='';
  if CDSLoteamento.Active Then
  Begin
     CDSLoteamento.First;
     if CDSLoteamento.RecordCount>0 then
     begin
       while not CDSLoteamento.Eof do begin
         varlotea := varlotea +quotedstr(CDSLoteamentoidloteamento.Text)+',';
         CDSLoteamento.Next;
       end;
       varlotea := copy(varlotea,1,length(varlotea)-1);
       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and idloteamento in ('+varlotea+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento,ordem,idloteamento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento,ordem,idloteamento');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      mensagem('Sem Filtro Definido.');
  percent.SetFocus;
end;

procedure TFrm_ReajusteDeParcelas2.DBGRecebDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if ZQRecebimento.FieldByName('saldo').Value <=0 Then
    DBGReceb.Canvas.Brush.Color:= $00FFEEDD; // coloque aqui a cor desejada
  if (ZQRecebimento.FieldByName('saldo').Value >0) and (ZQRecebimento.FieldByName('Dt_Vencimento').Value<date) Then
    DBGReceb.Canvas.Brush.Color:= $00DDDDFF; // coloque aqui a cor desejada
  if DBEIDReceb.Text = ZQRecebimento.FieldByName('idrecebimento').Text Then Begin
    DBGReceb.Canvas.Brush.Color := $006CFFFF;
    DBGReceb.Canvas.Font.Color := $00A80000;
//    DBGReceb.Canvas.Font.Style := [FsBold];
  end;
  DBGReceb.DefaultDrawDataCell(Rect, DBGReceb.columns[datacol].field, State);
  //Remove barra Horizontal}
 ShowScrollBar(DBGReceb.Handle,SB_HORZ,False);

{Remove barra Vertical}
 ShowScrollBar(DBGReceb.Handle,SB_VERT,False);

end;

procedure TFrm_ReajusteDeParcelas2.dxButton1Click(Sender: TObject);
var
vr,sld,corrigido,corrigido2:double;
ql,ms,varquadra:string;
posi,posi2:integer;
data:Tdatetime;
begin
  // adiciona 12 meses
  data:=IncMonth(strtodate('01/'+lmes.caption+'/'+xano.text),12);
  vr:=0;
  sld:=0;
  corrigido:=0;
  corrigido2:=0;
  if (percent.Value>0) and (ZQRecebtmp2.RecordCount>0) then
  begin
    dxButton4.Enabled:=false;
    dxButton5.Enabled:=false;

    dxButton1.Enabled:=false;

    ZQRecebtmp.First;
    barra2.Position:=0;
    barra2.Max:=ZQRecebtmp.RecordCount;
    while not ZQRecebtmp.Eof do
    begin
      barra2.Position:=ZQRecebtmp.RecNo;
      if  ZQRecebtmpmarcar.Value = '0' Then
      begin
        vr:=ZQRecebtmpvalor.Value;
        sld:=ZQRecebtmpsaldo.Value;
        corrigido:=vr+((vr*percent.Value)/100);
        corrigido2:=sld+((sld*percent.Value)/100);
        ZQRecebtmp.Edit;
        ZQRecebtmpsaldo_corr.Value:=ExRound(corrigido,2);
        ZQRecebtmp.post;
      end;
      ZQRecebtmp.Next;
    end;



    ZQRecebtmp2.First;
    barra2.Position:=0;
    barra2.Max:=ZQRecebtmp2.RecordCount;
    while not ZQRecebtmp2.Eof do
    begin
      barra2.Position:=ZQRecebtmp2.RecNo;
      if  ZQRecebtmp2marcar.Value = '0' Then
      begin
        vr:=ZQRecebtmp2valor.Value;
        sld:=ZQRecebtmp2saldo.Value;
        corrigido:=vr+((vr*percent.Value)/100);
        corrigido2:=sld+((sld*percent.Value)/100);
        ZQRecebtmp2.Edit;
        ZQRecebtmp2saldo_corr.Value:=ExRound(corrigido,2);
        ZQRecebtmp2.post;
      end;
      ZQRecebtmp2.Next;
    end;
    if simnao('Confirma o Reajuste das Parcelas ?','SIM') then
    begin
      vr:=0;
      sld:=0;
      corrigido:=0;
      corrigido2:=0;

      varquadra := '';
      ZQRecebtmp2.First;
      barra2.Position:=0;
      barra2.Max:=ZQRecebtmp2.RecordCount;

      if ZQRecebtmp2.RecordCount>0 then
      begin
        while not ZQRecebtmp2.Eof do
        begin
          barra2.Position:=ZQRecebtmp2.RecNo;        
          if  ZQRecebtmp2marcar.Value = '0' Then
          begin
            varquadra := varquadra +quotedstr(ZQRecebtmp2idrecebimento.Text)+',';
          end;  
          ZQRecebtmp2.Next;
        end;
        barra2.Position:=0;
        ZQRecebtmp2.First;
//        ZQRecebtmp.IndexName:='vencimento';
        ZQRecebtmp2.First;
        varquadra := copy(varquadra,1,length(varquadra)-1);
        DBGReceb.SelectedRows.Clear;
//        while not ZQRecebtmp.eof do
//        begin
          ZQRecebimento.close;
          ZQRecebimento.SQL.clear;
          if CBDesconsidera.Checked=false then
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento')
          else
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento');
          ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
          ZQRecebimento.Open;
          ZQRecebimento.First;
          barra2.Position:=0;
          barra2.Max:=ZQRecebimento.RecordCount;
          ZQRecebimento.DisableControls;
          while not ZQRecebimento.Eof do
          begin
            barra2.Position:=ZQRecebimento.RecNo;
            vr:=ZQRecebimento.FieldByName('Valor').Value;
            sld:=ZQRecebimento.FieldByName('saldo').Value;
            corrigido:=vr+((vr*percent.Value)/100);
            corrigido2:=sld+((sld*percent.Value)/100);
            ZQRecebimento.Edit;
            ZQRecebimento.FieldByName('Valor').Value:=ExRound(corrigido,2);
            ZQRecebimento.FieldByName('saldo').Value:=ExRound(corrigido2,2);
            ZQRecebimento.FieldByName('Reajustado').Value:='SIM';
            ZQRecebimento.FieldByName('Data_reajuste').Value:=date;
            ZQRecebimento.FieldByName('Proximo_Reajuste').Value:=copy(datetostr(data),4,7);
            ZQRecebimento.post;
            ZQRecebimento.Next;
          end;
          ZQRecebimento.Refresh;
        //  ZQRecebtmp.next;
        //end;
        ms:=UltimoDiaDoMes(lmes.caption+xano.text);

        ZQRecebimento.close;
        ZQRecebimento.SQL.clear;
{        ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
        ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
        ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
        if CBDesconsidera.Checked=false then
           ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
        else
           ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');

//        ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
        ZQRecebimento.Open;
        ZQRecebimento.First;
      end;
    end;
  end
  else
  begin
    mensagem('Informe o Percentual; Adicione os Títulos Selecionados no Grid !');
    percent.SetFocus;
    exit;
  end;


  ZQRecebtmp2.First;
  ql:=ZQRecebtmp2quadralote.Value;
  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);

  ZQEmpree.close;
  ZQEmpree.SQL.Clear;
  ZQEmpree.SQL.Add('Select * from imovel where quadra='+quotedstr(copy(ql,1,posi))+' and lote='+quotedstr(copy(ql,posi2,3)));
  ZQEmpree.open;
  ZQVenda.open;
  ZQEntrada.Close;
  ZQEntrada.SQL.Clear;
  ZQEntrada.SQL.Add('Select *, sum(valor) as entrada from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-E-%''  group by quadralote  order by ordem');
  ZQEntrada.open;

  ZQParcela.Close;
  ZQParcela.SQL.Clear;
  ZQParcela.SQL.Add('Select *, sum(valor) as parcela from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-P-%''  group by quadralote  order by ordem');
  ZQParcela.open;
  if (ZQEntrada.FieldByName('entrada').Value>0) or (ZQParcela.FieldByName('parcela').Value>0) then
  begin
    ZQVenda.Edit;
    ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
    ZQVenda.Post;
  end;
  barra2.Position:=0;
  barra2.Max:=ZQRecebtmp2.RecordCount;
  while not ZQRecebtmp2.Eof do
  begin
    barra2.Position:=ZQRecebtmp2.RecNo;
    if ql=ZQRecebtmp2quadralote.Value then
    begin
      ZQRecebtmp2.Next;
    end
    else
    begin
      ql:=ZQRecebtmp2quadralote.Value;
      posi:=pos('-',ql);
      dec(posi);

      posi2:=pos('-',ql);
      inc(posi2);

      ZQEmpree.close;
      ZQEmpree.SQL.Clear;
      ZQEmpree.SQL.Add('Select * from imovel where quadra='+quotedstr(copy(ql,1,posi))+' and lote='+quotedstr(copy(ql,posi2,3)));
      ZQEmpree.open;
      ZQVenda.open;
      ZQEntrada.Close;
      ZQEntrada.SQL.Clear;
      ZQEntrada.SQL.Add('Select *, sum(valor) as entrada from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-E-%''  order by ordem');
      ZQEntrada.open;

      ZQParcela.Close;
      ZQParcela.SQL.Clear;
      ZQParcela.SQL.Add('Select *, sum(valor) as parcela from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-P-%''  order by ordem');
      ZQParcela.open;

      if (ZQEntrada.FieldByName('entrada').Value>0) or (ZQParcela.FieldByName('parcela').Value>0) then
      begin
        ZQVenda.Edit;
        ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
        ZQVenda.Post;
      end;
    end;
  end;
  ql:=ZQRecebtmp2quadralote.Value;
  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);

  ZQEmpree.close;
  ZQEmpree.SQL.Clear;
  ZQEmpree.SQL.Add('Select * from imovel where quadra='+quotedstr(copy(ql,1,posi))+' and lote='+quotedstr(copy(ql,posi2,3)));
  ZQEmpree.open;
  ZQVenda.open;
  ZQEntrada.Close;
  ZQEntrada.SQL.Clear;
  ZQEntrada.SQL.Add('Select *, sum(valor) as entrada from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-E-%''  order by ordem');
  ZQEntrada.open;

  ZQParcela.Close;
  ZQParcela.SQL.Clear;
  ZQParcela.SQL.Add('Select *, sum(valor) as parcela from recebimento where quadralote='+quotedstr(ql)+' and documento like  ''%-P-%''  order by ordem');
  ZQParcela.open;
  if (ZQEntrada.FieldByName('entrada').Value>0) or (ZQParcela.FieldByName('parcela').Value>0) then
  begin
    ZQVenda.Edit;
    ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
    ZQVenda.Post;
  end;
  barra2.Position:=0;


  ZQEmpree.close;
  ZQVenda.close;
  ZQEntrada.close;
  ZQParcela.close;
  ql:='';
  // até aqui

  dxButton13.Enabled:=false;
  dxButton14.Enabled:=false;
  dxButton3.Enabled:=false;
  dxButton1.Enabled:=false;
  ZQRecebimento.EnableControls;
  barra2.Position:=0;
  ZQRecebtmp.IndexFieldNames:='adversanome';
  if ZQRecebtmp.RecordCount>0 then
  begin
    if FrmRelReajuste=nil then
       FrmRelReajuste:=TFrmRelReajuste.Create(Application);
    FrmRelReajuste.RLLabel2.Caption := 'Contratos Reajustados do mês/ano: '+xmes.Text+'/'+xano.Text;
    FrmRelReajuste.RLReport1.Preview;
  end;
    ZQRecebtmp.IndexFieldNames:='Dt_Vencimento';
end;

procedure TFrm_ReajusteDeParcelas2.xmesExit(Sender: TObject);
var
mmes,ms,xdata:string;
ds:Tdatetime;

begin
  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Março' then
     xdata:='03';
  if mmes='Abril' then
     xdata:='04';
  if mmes='Maio' then
     xdata:='05';
  if mmes='Junho' then
     xdata:='06';
  if mmes='Julho' then
     xdata:='07';
  if mmes='Agosto' then
     xdata:='08';
  if mmes='Setembro' then
     xdata:='09';
  if mmes='Outubro' then
     xdata:='10';
  if mmes='Novembro' then
     xdata:='11';
  if mmes='Dezembro' then
     xdata:='12';
  lmes.Caption:=xdata;

  ms:=UltimoDiaDoMes(xdata+xano.text);

  ZQRecebimento.close;
  ZQRecebimento.SQL.clear;
{  ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
  ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
  ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

  if CBDesconsidera.Checked=false then
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');

//  ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');  
  ZQRecebimento.Open;
  ZQRecebimento.First;
end;
procedure TFrm_ReajusteDeParcelas2.dxButton13Click(Sender: TObject);
begin
  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;
  ZQRecebimento.First;
end;

procedure TFrm_ReajusteDeParcelas2.dxButton14Click(Sender: TObject);
var
vlLinha,t: Integer;
begin
  Frm_ReajusteDeParcelas2.barra2.Position:=0;

  with DBGreceb.DataSource.DataSet do
  begin
    First;
    t:=RecordCount - 1;
    Frm_ReajusteDeParcelas2.barra2.Max:=t;
    for vlLinha := 0 to RecordCount - 1 do
    begin
     Frm_ReajusteDeParcelas2.barra2.Position:=vlLinha;
     DBGreceb.SelectedRows.CurrentRowSelected := True;
     Next;
    end;
  end;
  DBGreceb.SelectedRows.Refresh;
  ZQRecebimento.First;
  Frm_ReajusteDeParcelas2.barra2.Position:=0;
end;


procedure TFrm_ReajusteDeParcelas2.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if ZQRecebtmpsaldo.Value <=0 Then
    DBGrid2.Canvas.Brush.Color:= $00FFEEDD; // coloque aqui a cor desejada
  if (ZQRecebtmpsaldo.Value >0) and (ZQRecebtmpDt_Vencimento.Value<date) Then
    DBGrid2.Canvas.Brush.Color:= $00DDDDFF; // coloque aqui a cor desejada
  if DBid.Text = ZQRecebtmpidrecebimento.Text Then Begin
    DBGrid2.Canvas.Brush.Color := $006CFFFF;
    DBGrid2.Canvas.Font.Color := $00A80000;
//    DBGrid2.Canvas.Font.Style := [FsBold];
  end;
  DBGrid2.DefaultDrawDataCell(Rect, DBGrid2.columns[datacol].field, State);
  if Column.Field = ZQRecebtmpmarcar then
  begin
    DBGrid2.Canvas.FillRect(Rect);
    if ZQRecebtmpmarcar.Value = '0' then
      ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGrid2.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
    //Remove barra Horizontal}
 ShowScrollBar(DBGrid2.Handle,SB_HORZ,False);

end;

procedure TFrm_ReajusteDeParcelas2.dxButton3Click(Sender: TObject);
var
ms:string;
begin
  Frm_ReajusteDeParcelas2.barra2.Position:=0;
  ZQRecebtmp.close;
  ZQRecebtmp.createdataset;
  ZQRecebtmp2.close;
  ZQRecebtmp2.createdataset;

  GetSelectedRecord(ZQRecebimento, 'idrecebimento',DBGReceb.SelectedRows, ListBox1.Items);
  dxButton1.Enabled:=true;
  dxButton4.Enabled:=true;
  dxButton5.Enabled:=true;

  ms:=UltimoDiaDoMes(lmes.caption+xano.text);

  ZQRecebtmp.first;
  ZQRecebtmp.Filtered:=false;
  ZQRecebtmp.Filter:='Dt_Vencimento >= '+quotedstr('01/'+lmes.caption+'/'+xano.text)+ ' and Dt_Vencimento <= '+quotedstr(ms+'/'+lmes.caption+'/'+xano.text);
  ZQRecebtmp.Filtered:=true;
  ZQRecebtmp.first;
  ZQRecebtmp.recordcount;
    
  ZQRecebtmp2.first;
  ZQRecebtmp2.recordcount;
end;

procedure TFrm_ReajusteDeParcelas2.EloteamentoExit(Sender: TObject);
begin
  IF NOT EMPTY(Eloteamento.Text) Then Begin
    if Acha_loteamento(Eloteamento.Text) then Begin
      CDSLoteamento.Insert;
      CDSLoteamentoapelido.Value := DM_Tabelas.ZQAchaLotea.FieldByName('apelido').Value;
      CDSLoteamentoidloteamento.Value:=DM_Tabelas.ZQAchaLotea.FieldByName('idloteamento').Value;
      CDSLoteamento.Post;
    end;
    Eloteamento.Text := '';
    Eloteamento.SetFocus;
  end;
end;

procedure TFrm_ReajusteDeParcelas2.GroupBox5Enter(Sender: TObject);
var
mes_anos:string;
varquadra,vardevedor,varlotea:string;
mmes,ms,xdata:string;
ds:Tdatetime;
begin
  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Março' then
     xdata:='03';
  if mmes='Abril' then
     xdata:='04';
  if mmes='Maio' then
     xdata:='05';
  if mmes='Junho' then
     xdata:='06';
  if mmes='Julho' then
     xdata:='07';
  if mmes='Agosto' then
     xdata:='08';
  if mmes='Setembro' then
     xdata:='09';
  if mmes='Outubro' then
     xdata:='10';
  if mmes='Novembro' then
     xdata:='11';
  if mmes='Dezembro' then
     xdata:='12';
  lmes.Caption:=xdata;

  ms:=UltimoDiaDoMes(xdata+xano.text);




  mes_anos:=MesAno(strtodate('01/'+xdata+'/'+xano.text));
  Frm_AchaIgpm.ECidade.Text := mes_anos;
  DM_tabelas.ZQAchaIgpm.open;
  Frm_AchaIgpm.ShowModal;
  percent.Value:=DM_tabelas.ZQAchaIgpm.FieldByName('percentual').Value;
  DM_tabelas.ZQAchaIgpm.close;
end;

procedure TFrm_ReajusteDeParcelas2.percentExit(Sender: TObject);
begin
  dxbutton3.Enabled:=true;
  dxButton13.Enabled:=true;
  dxButton14.Enabled:=true;
  dxButton14.SetFocus;
end;

procedure TFrm_ReajusteDeParcelas2.DBGrid3KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSLoteamento.RecordCount>0) Then Begin
     CDSLoteamento.Delete;
  end;
end;

procedure TFrm_ReajusteDeParcelas2.DBGrid2MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ql:string;
  id:string;
begin
  if DBGrid2.SelectedField.fieldname = 'marcar' Then
  Begin
    ql:=ZQRecebtmpquadralote.Value;
    id:=ZQRecebtmpquadralote.value;
    ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);
    barra2.Max:=ZQRecebtmp2.RecordCount;
    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      barra2.Position:= ZQRecebtmp2.RecNo;
      ZQRecebtmp2.Edit;
      if  ZQRecebtmp2marcar.Value = '0' Then
      begin
        ZQRecebtmp2marcar.Value := '1';
      end
      else
      begin
        ZQRecebtmp2marcar.Value := '0';
      end;
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;
    ZQRecebtmp.Edit;
    if  ZQRecebtmpmarcar.Value = '0' Then
    begin
        ZQRecebtmpmarcar.Value := '1';
    end
    else
    begin
        ZQRecebtmpmarcar.Value := '0';
    end;
    ZQRecebtmp.post;
  end;
  barra2.Position:=0;
end;

procedure TFrm_ReajusteDeParcelas2.dxButton4Click(Sender: TObject);
var
  id,ql:string;
begin
  ZQRecebtmp.First;
  barra2.Position:=0;
  barra2.Max:=ZQRecebtmp.RecordCount;
  while not ZQRecebtmp.Eof do
  begin
    barra2.Position:=ZQRecebtmp.RecNo;
    ZQRecebtmp.Edit;
    ZQRecebtmpmarcar.Value := '0';
    ZQRecebtmp.post;

    ql:=ZQRecebtmpquadralote.Value;
    id:=ZQRecebtmpquadralote.value;
    ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);
    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      ZQRecebtmp2.Edit;
      ZQRecebtmp2marcar.Value := '0';
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;

    ZQRecebtmp.Next;
  end;
  ql:=ZQRecebtmpquadralote.Value;
  id:=ZQRecebtmpquadralote.value;
  ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);

  while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
  begin
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '0';
    ZQRecebtmp2.post;
    ZQRecebtmp2.Next;
  end;
  ZQRecebtmp.First;
  barra2.Position:=0;
end;

procedure TFrm_ReajusteDeParcelas2.dxButton5Click(Sender: TObject);
var
  id,ql:string;
begin
  ZQRecebtmp.First;
  barra2.Position:=0;
  barra2.Max:=ZQRecebtmp.RecordCount;
  while not ZQRecebtmp.Eof do
  begin
    barra2.Position:=ZQRecebtmp.RecNo;
    ZQRecebtmp.Edit;
    ZQRecebtmpmarcar.Value := '1';
    ZQRecebtmp.post;

    ql:=ZQRecebtmpquadralote.Value;

    id:=ZQRecebtmpquadralote.value;
    ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);

    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      ZQRecebtmp2.Edit;
      ZQRecebtmp2marcar.Value := '1';
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;

    ZQRecebtmp.Next;
  end;
  ql:=ZQRecebtmpquadralote.Value;
  id:=ZQRecebtmpquadralote.value;
  ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);
  while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
  begin
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '1';
    ZQRecebtmp2.post;
    ZQRecebtmp2.Next;
  end;
  ZQRecebtmp.First;
  barra2.Position:=0;
end;

procedure TFrm_ReajusteDeParcelas2.DBGrid5DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if ZQRecebtmp2saldo.Value <=0 Then
    DBGrid5.Canvas.Brush.Color:= $00FFEEDD; // coloque aqui a cor desejada
  if (ZQRecebtmp2saldo.Value >0) and (ZQRecebtmp2Dt_Vencimento.Value<date) Then
  begin
    DBGrid5.Canvas.Brush.Color:= $00DDDDFF; // coloque aqui a cor desejada
  end;
  DBGrid5.DefaultDrawDataCell(Rect, DBGrid5.columns[datacol].field, State);
  if Column.Field = ZQRecebtmp2marcar then
  begin
    DBGrid5.Canvas.FillRect(Rect);
    if ZQRecebtmp2marcar.Value = '0' then
      ImageList1.Draw(DBGrid5.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGrid5.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
    //Remove barra Horizontal}
 ShowScrollBar(DBGrid5.Handle,SB_HORZ,False);


end;

procedure TFrm_ReajusteDeParcelas2.DBGrid5MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  var
  ql:string;
begin
  if DBGrid5.SelectedField.fieldname = 'marcar' Then
  Begin
    ql:=ZQRecebtmp2quadralote.Value;
    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      ZQRecebtmp2.Edit;
      if  ZQRecebtmp2marcar.Value = '0' Then
      begin
        ZQRecebtmp2marcar.Value := '1';
      end
      else
      begin
        ZQRecebtmp2marcar.Value := '0';
      end;
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;
  end;
  barra2.Position:=0;

end;

procedure TFrm_ReajusteDeParcelas2.cbconferenciaClick(Sender: TObject);
begin
  if cbconferencia.Checked=true then
     DBGrid5.Visible:=true
  else
    DBGrid5.Visible:=false;   
end;


procedure TFrm_ReajusteDeParcelas2.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1quadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1lote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1lado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradasomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParceladocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas2, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreelote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreelado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreematri', 'matri', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas2, 'ZQEmpree', 'ZQEmpreeApelido', 'Apelido', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'loteamento_idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);

end.
