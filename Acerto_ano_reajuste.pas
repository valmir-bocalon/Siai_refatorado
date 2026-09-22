unit Acerto_ano_reajuste;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, XNum, XLabel3D, Mask, XDate, StdCtrls, Grids, DBGrids, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBCtrls, TFlatGaugeUnit,DbiProcs,Shellapi, ImgList, ExtCtrls,
  ComCtrls, System.ImageList;

type
  TFrm_Acerto_ano = class(TForm)
    XBanner4: TXBanner;
    GBAdversa: TGroupBox;
    XBanner10: TXBanner;
    XBanner6: TXBanner;
    DBGrid4: TDBGrid;
    XLabel3D1: TXLabel3D;
    DXBFechar: TdxButtonArround;
    dxButton1: TdxButtonArround;
    DS_quadraLote: TDataSource;
    DS_Adversatemp: TDataSource;
    ZQAdvsersatemp: TZQuery;

































    CDSAdversaTemp2: TClientDataSet;
    CDSAdversaTemp2CodAdversa: TIntegerField;
    CDSAdversaTemp2nomeadversa: TWideStringField;
    DataCDSAdversaTEmp2: TDataSource;
    CDSQuadraLote: TClientDataSet;
    CDSQuadraLoteQuadra: TWideStringField;
    CDSQuadraLotelote: TWideStringField;
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
    Eadversa: TEdit;
    DBGReceb: TDBGrid;
    DXBPrimeiro: TdxButtonArround;
    DXBAnterior: TdxButtonArround;
    DXBProximo: TdxButtonArround;
    DXBUltimo: TdxButtonArround;
    DXBPesquisar: TdxButtonArround;
    dxButton2: TdxButtonArround;
    DBEIDReceb: TDBEdit;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    xmes: TComboBox;
    xano: TComboBox;
    lmes: TLabel;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;

































    DBGrid2: TDBGrid;
    DataZQRecebtmp: TDataSource;
    DBid: TDBEdit;
    ListBox1: TListBox;


    dxButton3: TdxButtonArround;
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
    CDSLoteamentonomeloteamento: TWideStringField;
    CDSLoteamentocodcontabancaria: TIntegerField;
    CDSLoteamentoapelido: TWideStringField;
    GroupBox3: TGroupBox;
    XBanner2: TXBanner;
    XBanner5: TXBanner;
    DBGrid3: TDBGrid;
    Eloteamento: TEdit;
    dxButton14: TdxButtonArround;
    dxButton13: TdxButtonArround;


    CBDesconsidera: TCheckBox;
    ImageList1: TImageList;
    dxButton4: TdxButtonArround;
    dxButton5: TdxButtonArround;
    DBGrid5: TDBGrid;
    DataZQRecebtmp2: TDataSource;
    barra2: TProgressBar;
    cbconferencia: TCheckBox;

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
    ZQRecebtmp: TClientDataSet;
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
    ZQRecebtmpproximo_reajuste: TStringField;
    ZQRecebtmpParcelas_fixas: TStringField;
    destadata: TMaskEdit;
    Label1: TLabel;
    Panel1: TPanel;
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

    procedure PrepararConsultaRecebimentoAntesDeAbrir(DataSet: TDataSet);
    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Acerto_ano: TFrm_Acerto_ano;

implementation

uses tabelas, PesqRecebimento,funcoes, AchaIgpm, PesqRecebimento2,
  uRuntimeFields, uRecebimentoNomes;

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
      Frm_Acerto_ano.barra2.Position:=0;
      t:=BookMarkList.Count -1;
      Frm_Acerto_ano.barra2.Max:=t;
      for i := 0 to BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Frm_Acerto_ano.barra2.Position:=i;
        Strings.Add(DataSet.FindField(FieldName).AsString);
        Frm_Acerto_ano.ZQRecebimento.Filtered:=false;
        Frm_Acerto_ano.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        Frm_Acerto_ano.ZQRecebimento.Filtered:=true;
        Frm_Acerto_ano.ZQRecebimento.First;
        if Frm_Acerto_ano.ZQRecebimento.recordcount>0 then
        begin
  //        Frm_Acerto_ano.barra.MaxValue:=Frm_Acerto_ano.ZQRecebimento.recordcount;
          while not Frm_Acerto_ano.ZQRecebimento.Eof do
          begin
//            Frm_Acerto_ano.barra.Progress:=Frm_Acerto_ano.ZQRecebimento.recno;
            Frm_Acerto_ano.ZQRecebtmp.Insert;
            Frm_Acerto_ano.ZQRecebtmpidrecebimento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            Frm_Acerto_ano.ZQRecebtmpdocumento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('documento').AsString;
            Frm_Acerto_ano.ZQRecebtmpcliente.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('cliente').AsInteger;
            Frm_Acerto_ano.ZQRecebtmpusuario.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('usuario').AsInteger;
            Frm_Acerto_ano.ZQRecebtmpDt_Entrada.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
            Frm_Acerto_ano.ZQRecebtmpDt_Vencimento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
            Frm_Acerto_ano.ZQRecebtmpValor.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Valor').AsFloat;
            Frm_Acerto_ano.ZQRecebtmpVrDoc.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('VrDoc').AsFloat;
            Frm_Acerto_ano.ZQRecebtmpordem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('ordem').AsString;
            Frm_Acerto_ano.ZQRecebtmpTipDoc.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('TipDoc').AsString;
            Frm_Acerto_ano.ZQRecebtmpsaldo.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('saldo').AsFloat;
            Frm_Acerto_ano.ZQRecebtmporigem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('origem').AsString;
            Frm_Acerto_ano.ZQRecebtmpadversa.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('adversa').AsInteger;
            Frm_Acerto_ano.ZQRecebtmprecpag.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('recpag').AsString;
            Frm_Acerto_ano.ZQRecebtmpnumordem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('numordem').AsInteger;
            Frm_Acerto_ano.ZQRecebtmpadversanome.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('adversanome').AsString;
            Frm_Acerto_ano.ZQRecebtmpidloteamento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('idloteamento').AsInteger;
            Frm_Acerto_ano.ZQRecebtmpvenda_idvenda.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('venda_idvenda').AsInteger;
            Frm_Acerto_ano.ZQRecebtmpquadralote.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('quadralote').AsString;
            Frm_Acerto_ano.ZQRecebtmpnomeadversa.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('nomeadversa').AsString;
            Frm_Acerto_ano.ZQRecebtmpParcelas_fixas.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Parcelas_fixas').AsString;
            Frm_Acerto_ano.ZQRecebtmpproximo_reajuste.Value:=Frm_Acerto_ano.lmes.caption+'/'+Frm_Acerto_ano.xano.text;
            Frm_Acerto_ano.ZQRecebtmpmarcar.Value := '0';
            Frm_Acerto_ano.ZQRecebtmp.Post;




            Frm_Acerto_ano.ZQRecebtmp2.Insert;
            Frm_Acerto_ano.ZQRecebtmp2idrecebimento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            Frm_Acerto_ano.ZQRecebtmp2documento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('documento').AsString;
            Frm_Acerto_ano.ZQRecebtmp2cliente.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('cliente').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2usuario.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('usuario').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2Dt_Entrada.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
            Frm_Acerto_ano.ZQRecebtmp2Dt_Vencimento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
            Frm_Acerto_ano.ZQRecebtmp2Valor.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Valor').AsFloat;
            Frm_Acerto_ano.ZQRecebtmp2VrDoc.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('VrDoc').AsFloat;
            Frm_Acerto_ano.ZQRecebtmp2ordem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('ordem').AsString;
            Frm_Acerto_ano.ZQRecebtmp2TipDoc.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('TipDoc').AsString;
            Frm_Acerto_ano.ZQRecebtmp2saldo.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('saldo').AsFloat;
            Frm_Acerto_ano.ZQRecebtmp2origem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('origem').AsString;
            Frm_Acerto_ano.ZQRecebtmp2adversa.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('adversa').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2recpag.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('recpag').AsString;
            Frm_Acerto_ano.ZQRecebtmp2numordem.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('numordem').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2adversanome.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('adversanome').AsString;
            Frm_Acerto_ano.ZQRecebtmp2idloteamento.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('idloteamento').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2venda_idvenda.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('venda_idvenda').AsInteger;
            Frm_Acerto_ano.ZQRecebtmp2quadralote.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('quadralote').AsString;
            Frm_Acerto_ano.ZQRecebtmp2nomeadversa.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('nomeadversa').AsString;
            Frm_Acerto_ano.ZQRecebtmp2Parcelas_fixas.Value:=Frm_Acerto_ano.ZQRecebimento.FieldByName('Parcelas_fixas').AsString;
            Frm_Acerto_ano.ZQRecebtmp2proximo_reajuste.Value:=Frm_Acerto_ano.lmes.caption+'/'+Frm_Acerto_ano.xano.text;
            Frm_Acerto_ano.ZQRecebtmp2marcar.Value := '0';            
            Frm_Acerto_ano.ZQRecebtmp2.Post;



            Frm_Acerto_ano.ZQRecebimento.Next;
          end;
        end;
        Frm_Acerto_ano.ZQRecebimento.Filtered:=false;
        Frm_Acerto_ano.barra2.Position:=0;
      end;
    finally
      GotoBookMark(bm);
      FreeBookMark(bm);
      Frm_Acerto_ano.barra2.Position:=0;
      EnableControls;
    end;
  end; // With DataSet do
end;


procedure TFrm_Acerto_ano.DXBFecharClick(Sender: TObject);
begin
  close;
end;
procedure TFrm_Acerto_ano.EadversaExit(Sender: TObject);
begin
  IF NOT EMPTY(Eadversa.Text) Then Begin
    if AchaCliente(400, 300, Eadversa.Text) then Begin
      CDSAdversaTemp2.Insert;
      CDSAdversaTemp2CodAdversa.Value := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      CDSAdversaTemp2.Post;
    end;
    Eadversa.Text := '';
    Eadversa.SetFocus;      
  end;

end;

procedure TFrm_Acerto_ano.EloteExit(Sender: TObject);
begin
  if not empty(Elote.Text+EQuadra.Text) Then Begin
    ZQVen1.SQL.Clear;
    ZQVen1.SQL.Add('Select * from venda join imovel on idimovel=imovel');
    ZQVen1.SQL.Add(' where quadra='+quotedstr(Equadra.Text)+' and lote='+quotedstr(Elote.Text));
    ZQVen1.Open;
    if ZQVen1.RecordCount>0 Then Begin
      CDSQuadraLote.Insert;
      CDSQuadraLoteidlote.Value := ZQVen1.FieldByName('idimovel').AsLargeInt;
      CDSQuadraLoteQuadra.Value := Equadra.Text+'-'+Elote.Text;
      CDSQuadraLote.Post;
    end
    else
      Showmessage('Esta Quadra e lote não foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end;
end;

procedure TFrm_Acerto_ano.FormShow(Sender: TObject);
var
mmes,ms,xdata:string;
ds:Tdatetime;

begin
  { Os campos de comprador e loteamento sao calculados por lookup. Abra as
    fontes antes de abrir a consulta principal para o primeiro registro. }
  if not DM_Tabelas.ZqParticipante.Active then
    DM_Tabelas.ZqParticipante.Open;
  if not DM_Tabelas.ZQLoteamento.Active then
    DM_Tabelas.ZQLoteamento.Open;

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
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') order by DT_Vencimento');

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
procedure TFrm_Acerto_ano.botoes_setas;
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

procedure TFrm_Acerto_ano.DXBPrimeiroClick(Sender: TObject);
begin
  ZQRecebimento.First;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_ano.DXBAnteriorClick(Sender: TObject);
begin
  ZQRecebimento.Prior;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_ano.DXBProximoClick(Sender: TObject);
begin
  ZQRecebimento.Next;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_ano.DXBUltimoClick(Sender: TObject);
begin
  ZQRecebimento.Last;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_ano.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento2.Top := Frm_Acerto_ano.Top+100;
  FrmPesqRecebimento2.Left := Frm_Acerto_ano.Left;
  FrmPesqRecebimento2Target := Self;
  try
    AbrirModal(Self, FrmPesqRecebimento2);
  finally
    FrmPesqRecebimento2Target := nil;
  end;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_ano.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrm_Acerto_ano.FormClose(Sender: TObject;
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

procedure TFrm_Acerto_ano.DBGrid4KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;

end;

procedure TFrm_Acerto_ano.DBGrid1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;

end;

procedure TFrm_Acerto_ano.dxButton2Click(Sender: TObject);
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
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (numboleto is null or numboleto='''') and Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+'  order by DT_Vencimento,ordem,quadralote')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' order by DT_Vencimento,ordem,quadralote');
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
         ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') order by DT_Vencimento,ordem,adversa')
      else
         ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+')  order by DT_Vencimento,ordem,adversa');
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
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') order by DT_Vencimento,ordem,idloteamento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+')  order by DT_Vencimento,ordem,idloteamento');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      mensagem('Sem Filtro Definido.');
  dxbutton3.Enabled:=true;
  dxButton13.Enabled:=true;
  dxButton14.Enabled:=true;
  dxButton14.SetFocus;
end;

procedure TFrm_Acerto_ano.DBGRecebDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if ZQRecebimento.FieldByName('saldo').AsFloat <=0 Then
    DBGReceb.Canvas.Brush.Color:= $00FFEEDD; // coloque aqui a cor desejada
  if (ZQRecebimento.FieldByName('saldo').AsFloat >0) and (ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime<date) Then
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

procedure TFrm_Acerto_ano.dxButton1Click(Sender: TObject);
var
vr,sld,corrigido,corrigido2:double;
ql,ms,varquadra:string;
posi,posi2:integer;
data:Tdatetime;
begin
  if simnao('Confirma o processamento ?','SIM') then
  begin
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
          if not empty(charrem(destadata.Text)) then
          begin

            if CBDesconsidera.Checked=false then
               ZQRecebimento.SQL.Add(' Select * from Recebimento where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+')  order by DT_Vencimento')
            else
               ZQRecebimento.SQL.Add(' Select * from Recebimento where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+')  order by DT_Vencimento');
            ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
          end
          else
          begin
            if CBDesconsidera.Checked=false then
               ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+')  order by DT_Vencimento')
            else
               ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idrecebimento in ('+varquadra+')'+')  order by DT_Vencimento');

          end;
    //      ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
          ZQRecebimento.Open;
          ZQRecebimento.First;
          barra2.Position:=0;
          barra2.Max:=ZQRecebimento.RecordCount;
          ZQRecebimento.DisableControls;
          while not ZQRecebimento.Eof do
          begin
            barra2.Position:=ZQRecebimento.RecNo;
            ZQRecebimento.Edit;
            ZQRecebimento.FieldByName('Data_reajuste').AsDateTime:=strtodate('01/'+lmes.caption+'/'+xano.text);
            ZQRecebimento.FieldByName('Proximo_Reajuste').AsString:=lmes.caption+'/'+xano.text;
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
           ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
        else
           ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');

//        ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
        ZQRecebimento.Open;
        ZQRecebimento.First;
      end;
  end;


  ZQRecebtmp2.First;
  ql:=ZQRecebtmp2quadralote.Value;
  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);


  ql:=ZQRecebtmp2quadralote.Value;
  posi:=pos('-',ql);
  dec(posi);

  posi2:=pos('-',ql);
  inc(posi2);

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
end;

procedure TFrm_Acerto_ano.xmesExit(Sender: TObject);
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
     ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
  else
     ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');

//  ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');  
  ZQRecebimento.Open;
  ZQRecebimento.First;
end;
procedure TFrm_Acerto_ano.dxButton13Click(Sender: TObject);
begin
  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;
  ZQRecebimento.First;
end;

procedure TFrm_Acerto_ano.dxButton14Click(Sender: TObject);
var
vlLinha,t: Integer;
begin
  Frm_Acerto_ano.barra2.Position:=0;

  with DBGreceb.DataSource.DataSet do
  begin
    First;
    t:=RecordCount - 1;
    Frm_Acerto_ano.barra2.Max:=t;
    for vlLinha := 0 to RecordCount - 1 do
    begin
     Frm_Acerto_ano.barra2.Position:=vlLinha;
     DBGreceb.SelectedRows.CurrentRowSelected := True;
     Next;
    end;
  end;
  DBGreceb.SelectedRows.Refresh;
  ZQRecebimento.First;
  Frm_Acerto_ano.barra2.Position:=0;
end;


procedure TFrm_Acerto_ano.DBGrid2DrawColumnCell(Sender: TObject;
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

procedure TFrm_Acerto_ano.dxButton3Click(Sender: TObject);
var
ms:string;
begin
  Frm_Acerto_ano.barra2.Position:=0;
  Panel1.visible:=true;
  Application.ProcessMessages;
  ZQRecebtmp.close;
  ZQRecebtmp.createdataset;
  ZQRecebtmp2.close;
  ZQRecebtmp2.createdataset;

  GetSelectedRecord(ZQRecebimento, 'idrecebimento',DBGReceb.SelectedRows, ListBox1.Items);
  dxButton1.Enabled:=true;
  dxButton4.Enabled:=true;
  dxButton5.Enabled:=true;
  Panel1.visible:=false;
  Application.ProcessMessages;

  ms:=UltimoDiaDoMes(lmes.caption+xano.text);

  ZQRecebtmp.first;
{  ZQRecebtmp.Filtered:=false;
  ZQRecebtmp.Filter:='Dt_Vencimento >= '+quotedstr('01/'+lmes.caption+'/'+xano.text)+ ' and Dt_Vencimento <= '+quotedstr(ms+'/'+lmes.caption+'/'+xano.text);
  ZQRecebtmp.Filtered:=true;
  ZQRecebtmp.first;}
    
  ZQRecebtmp2.first;
end;

procedure TFrm_Acerto_ano.EloteamentoExit(Sender: TObject);
begin
  IF NOT EMPTY(Eloteamento.Text) Then Begin
    if Acha_loteamento(Eloteamento.Text) then Begin
      CDSLoteamento.Insert;
      CDSLoteamentoapelido.Value := DM_Tabelas.ZQAchaLotea.FieldByName('apelido').AsString;
      CDSLoteamentoidloteamento.Value:=DM_Tabelas.ZQAchaLotea.FieldByName('idloteamento').AsLargeInt;
      CDSLoteamento.Post;
    end;
    Eloteamento.Text := '';
    Eloteamento.SetFocus;
  end;
end;

procedure TFrm_Acerto_ano.DBGrid3KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSLoteamento.RecordCount>0) Then Begin
     CDSLoteamento.Delete;
  end;
end;

procedure TFrm_Acerto_ano.DBGrid2MouseUp(Sender: TObject;
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

procedure TFrm_Acerto_ano.dxButton4Click(Sender: TObject);
var
  id,ql:string;
begin
  ZQRecebtmp2.First;
  barra2.Position:=0;
  barra2.Max:=ZQRecebtmp2.RecordCount;
  while not ZQRecebtmp2.Eof do
  begin
    barra2.Position:=ZQRecebtmp2.RecNo;
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '0';
    ZQRecebtmp2.post;

{    ql:=ZQRecebtmpquadralote.Value;
    id:=ZQRecebtmpquadralote.value;
    ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);
    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      ZQRecebtmp2.Edit;
      ZQRecebtmp2marcar.Value := '0';
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;
 }
    ZQRecebtmp2.Next;
  end;
{  ql:=ZQRecebtmpquadralote.Value;
  id:=ZQRecebtmpquadralote.value;
  ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);

  while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
  begin
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '0';
    ZQRecebtmp2.post;
    ZQRecebtmp2.Next;
  end;             }
  ZQRecebtmp2.First;
  barra2.Position:=0;
end;

procedure TFrm_Acerto_ano.dxButton5Click(Sender: TObject);
var
  id,ql:string;
begin
  ZQRecebtmp2.First;
  barra2.Position:=0;
  barra2.Max:=ZQRecebtmp.RecordCount;
  while not ZQRecebtmp2.Eof do
  begin
    barra2.Position:=ZQRecebtmp2.RecNo;
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '1';
    ZQRecebtmp2.post;

{    ql:=ZQRecebtmpquadralote.Value;

    id:=ZQRecebtmpquadralote.value;
    ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);

    while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
    begin
      ZQRecebtmp2.Edit;
      ZQRecebtmp2marcar.Value := '1';
      ZQRecebtmp2.post;
      ZQRecebtmp2.Next;
    end;
}
    ZQRecebtmp2.Next;
  end;
{  ql:=ZQRecebtmpquadralote.Value;
  id:=ZQRecebtmpquadralote.value;
  ZQRecebtmp2.Locate('quadralote',id,[loCaseinsensitive]);
  while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
  begin
    ZQRecebtmp2.Edit;
    ZQRecebtmp2marcar.Value := '1';
    ZQRecebtmp2.post;
    ZQRecebtmp2.Next;
  end;
  ZQRecebtmp.First;}
  barra2.Position:=0;
end;

procedure TFrm_Acerto_ano.DBGrid5DrawColumnCell(Sender: TObject;
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

procedure TFrm_Acerto_ano.DBGrid5MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  var
  ql:string;
begin
  if DBGrid5.SelectedField.fieldname = 'marcar' Then
  Begin
//    ql:=ZQRecebtmp2quadralote.Value;
  //  while (ql=ZQRecebtmp2quadralote.Value) and (not ZQRecebtmp2.Eof) do
//    begin
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
  //    ZQRecebtmp2.Next;
//    end;
  end;
  barra2.Position:=0;

end;

procedure TFrm_Acerto_ano.cbconferenciaClick(Sender: TObject);
begin
  if cbconferencia.Checked=true then
     DBGrid5.Visible:=true
  else
    DBGrid5.Visible:=false;   
end;


procedure TFrm_Acerto_ano.PrepararConsultaRecebimentoAntesDeAbrir(
  DataSet: TDataSet);
begin
  PrepararConsultaRecebimentoComNomes(DataSet);
end;

procedure TFrm_Acerto_ano.AfterConstruction;
begin
  inherited AfterConstruction;
  ZQRecebimento.BeforeOpen := PrepararConsultaRecebimentoAntesDeAbrir;
  ZQTempReceber.BeforeOpen := PrepararConsultaRecebimentoAntesDeAbrir;
  PrepararConsultaRecebimentoComNomes(ZQRecebimento);
  PrepararConsultaRecebimentoComNomes(ZQTempReceber);
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_ano, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_ano, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
