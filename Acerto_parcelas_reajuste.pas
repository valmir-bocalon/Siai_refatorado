unit Acerto_parcelas_reajuste;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, XNum, XLabel3D, Mask, XDate, StdCtrls, Grids, DBGrids,
    dxButton, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBCtrls, TFlatGaugeUnit,DbiProcs,Shellapi, ImgList, ExtCtrls,
  ComCtrls, dxCore2, System.ImageList, FnpNumericEdit, dxCoreA;

type
  TFrm_Acerto_parcelas = class(TForm)
    XBanner4: TXBanner;
    GBAdversa: TGroupBox;
    XBanner10: TXBanner;
    XBanner6: TXBanner;
    DBGrid4: TDBGrid;
    XLabel3D1: TXLabel3D;
    DXBFechar: TdxButton;
    dxButton1: TdxButton;
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
    DXBPrimeiro: TdxButton;
    DXBAnterior: TdxButton;
    DXBProximo: TdxButton;
    DXBUltimo: TdxButton;
    DXBPesquisar: TdxButton;
    dxButton2: TdxButton;
    DBEIDReceb: TDBEdit;
    GroupBox1: TGroupBox;
    XBanner1: TXBanner;
    lmes: TLabel;
    DS_Recebimento: TDataSource;
    ZQRecebimento: TZQuery;

































    DBGrid2: TDBGrid;
    DataZQRecebtmp: TDataSource;
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
    CDSLoteamentonomeloteamento: TWideStringField;
    CDSLoteamentocodcontabancaria: TIntegerField;
    CDSLoteamentoapelido: TWideStringField;
    GroupBox3: TGroupBox;
    XBanner2: TXBanner;
    XBanner5: TXBanner;
    DBGrid3: TDBGrid;
    Eloteamento: TEdit;
    dxButton14: TdxButton;
    dxButton13: TdxButton;


    CBDesconsidera: TCheckBox;
    ImageList1: TImageList;
    dxButton4: TdxButton;
    dxButton5: TdxButton;
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
    xprnovo: TFnpNumericEdit;
    Label2: TLabel;
    Panel1: TPanel;
    Label4: TLabel;
    xreajuste: TMaskEdit;
    xvrnovo: TFnpNumericEdit;
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
  Frm_Acerto_parcelas: TFrm_Acerto_parcelas;

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
      Frm_Acerto_parcelas.barra2.Position:=0;
      t:=BookMarkList.Count -1;
      Frm_Acerto_parcelas.barra2.Max:=t;
      for i := 0 to BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Frm_Acerto_parcelas.barra2.Position:=i;
        Strings.Add(DataSet.FindField(FieldName).AsString);
        Frm_Acerto_parcelas.ZQRecebimento.Filtered:=false;
        Frm_Acerto_parcelas.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        Frm_Acerto_parcelas.ZQRecebimento.Filtered:=true;
        Frm_Acerto_parcelas.ZQRecebimento.First;
        if Frm_Acerto_parcelas.ZQRecebimento.recordcount>0 then
        begin
  //        Frm_Acerto_parcelas.barra.MaxValue:=Frm_Acerto_parcelas.ZQRecebimento.recordcount;
          while not Frm_Acerto_parcelas.ZQRecebimento.Eof do
          begin
//            Frm_Acerto_parcelas.barra.Progress:=Frm_Acerto_parcelas.ZQRecebimento.recno;
            Frm_Acerto_parcelas.ZQRecebtmp.Insert;
            Frm_Acerto_parcelas.ZQRecebtmpidrecebimento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            Frm_Acerto_parcelas.ZQRecebtmpdocumento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('documento').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpcliente.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('cliente').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmpusuario.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('usuario').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmpDt_Entrada.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
            Frm_Acerto_parcelas.ZQRecebtmpDt_Vencimento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
            Frm_Acerto_parcelas.ZQRecebtmpValor.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Valor').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmpVrDoc.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('VrDoc').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmpordem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('ordem').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpTipDoc.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('TipDoc').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpsaldo.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('saldo').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmporigem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('origem').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpadversa.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('adversa').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmprecpag.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('recpag').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpnumordem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('numordem').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmpadversanome.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('adversanome').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpidloteamento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('idloteamento').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmpvenda_idvenda.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('venda_idvenda').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmpquadralote.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('quadralote').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpnomeadversa.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('nomeadversa').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpParcelas_fixas.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Parcelas_fixas').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpproximo_reajuste.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Proximo_Reajuste').AsString;
            Frm_Acerto_parcelas.ZQRecebtmpmarcar.Value := '0';
            Frm_Acerto_parcelas.ZQRecebtmp.Post;




            Frm_Acerto_parcelas.ZQRecebtmp2.Insert;
            Frm_Acerto_parcelas.ZQRecebtmp2idrecebimento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
            Frm_Acerto_parcelas.ZQRecebtmp2documento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('documento').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2cliente.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('cliente').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2usuario.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('usuario').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2Dt_Entrada.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime;
            Frm_Acerto_parcelas.ZQRecebtmp2Dt_Vencimento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime;
            Frm_Acerto_parcelas.ZQRecebtmp2Valor.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Valor').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmp2VrDoc.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('VrDoc').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmp2ordem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('ordem').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2TipDoc.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('TipDoc').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2saldo.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('saldo').AsFloat;
            Frm_Acerto_parcelas.ZQRecebtmp2origem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('origem').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2adversa.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('adversa').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2recpag.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('recpag').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2numordem.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('numordem').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2adversanome.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('adversanome').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2idloteamento.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('idloteamento').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2venda_idvenda.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('venda_idvenda').AsInteger;
            Frm_Acerto_parcelas.ZQRecebtmp2quadralote.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('quadralote').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2nomeadversa.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('nomeadversa').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2Parcelas_fixas.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Parcelas_fixas').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2proximo_reajuste.Value:=Frm_Acerto_parcelas.ZQRecebimento.FieldByName('Proximo_Reajuste').AsString;
            Frm_Acerto_parcelas.ZQRecebtmp2marcar.Value := '0';
            Frm_Acerto_parcelas.ZQRecebtmp2.Post;



            Frm_Acerto_parcelas.ZQRecebimento.Next;
          end;
        end;
        Frm_Acerto_parcelas.ZQRecebimento.Filtered:=false;
        Frm_Acerto_parcelas.barra2.Position:=0;
      end;
    finally
      GotoBookMark(bm);
      FreeBookMark(bm);
      Frm_Acerto_parcelas.barra2.Position:=0;
      EnableControls;
    end;
  end; // With DataSet do
end;


procedure TFrm_Acerto_parcelas.DXBFecharClick(Sender: TObject);
begin
  close;
end;
procedure TFrm_Acerto_parcelas.EadversaExit(Sender: TObject);
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

procedure TFrm_Acerto_parcelas.EloteExit(Sender: TObject);
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

procedure TFrm_Acerto_parcelas.FormShow(Sender: TObject);
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
procedure TFrm_Acerto_parcelas.botoes_setas;
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

procedure TFrm_Acerto_parcelas.DXBPrimeiroClick(Sender: TObject);
begin
  ZQRecebimento.First;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_parcelas.DXBAnteriorClick(Sender: TObject);
begin
  ZQRecebimento.Prior;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_parcelas.DXBProximoClick(Sender: TObject);
begin
  ZQRecebimento.Next;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_parcelas.DXBUltimoClick(Sender: TObject);
begin
  ZQRecebimento.Last;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_parcelas.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento2.Top := Frm_Acerto_parcelas.Top+100;
  FrmPesqRecebimento2.Left := Frm_Acerto_parcelas.Left;
  FrmPesqRecebimento2Target := Self;
  try
    FrmPesqRecebimento2.showmodal;
  finally
    FrmPesqRecebimento2Target := nil;
  end;
  DBGReceb.SetFocus;

end;

procedure TFrm_Acerto_parcelas.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrm_Acerto_parcelas.FormClose(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.DBGrid4KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;

end;

procedure TFrm_Acerto_parcelas.DBGrid1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;

end;

procedure TFrm_Acerto_parcelas.dxButton2Click(Sender: TObject);
var
varquadra,vardevedor,varloteamento:string;
mmes,ms,xdata:string;
ds:Tdatetime;
begin
   varquadra := '';
  if (CDSQuadraLote.Active) and (CDSLoteamento.Active=false) Then
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

  if (CDSQuadraLote.Active) and (CDSLoteamento.Active) Then
  Begin
     CDSQuadraLote.First;
     if CDSQuadraLote.RecordCount>0 then
     begin
       while not CDSQuadraLote.Eof do begin
         varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
         CDSQuadraLote.Next;
       end;
       varquadra := copy(varquadra,1,length(varquadra)-1);

       CDSLoteamento.First;
       if CDSLoteamento.RecordCount>0 then
       begin
         while not CDSLoteamento.Eof do begin
           varloteamento := varloteamento +quotedstr(CDSLoteamentoidloteamento.Text)+',';
           CDSLoteamento.Next;
         end;
       end;
       varloteamento := copy(varloteamento,1,length(varloteamento)-1);
       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (numboleto is null or numboleto='''') and Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and idloteamento in ('+varloteamento+')'+' and quadralote in ('+varquadra+')'+'  order by DT_Vencimento,ordem,quadralote')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and idloteamento in ('+varloteamento+')'+' and quadralote in ('+varquadra+')'+' order by DT_Vencimento,ordem,quadralote');
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

{  varlotea:='';
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
{       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') order by DT_Vencimento,ordem,idloteamento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+')  order by DT_Vencimento,ordem,idloteamento');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;}
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      showmessage('Sem Filtro Definido.');
{  dxbutton3.Enabled:=true;
  dxButton13.Enabled:=true;
  dxButton14.Enabled:=true;
  dxButton14.SetFocus;}
  dxButton1.Enabled:=true;
  dxButton1.SetFocus;
end;

procedure TFrm_Acerto_parcelas.DBGRecebDrawColumnCell(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.dxButton1Click(Sender: TObject);
var
vr,sld,corrigido,corrigido2:double;
ql,ms,varquadra,varloteamento,perc:string;
posi,posi2:integer;
data:Tdatetime;
begin
  if simnao('Confirma o processamento ?','SIM') then
  begin
    varquadra := '';
    if xprnovo.Value>0 then
    begin
      xvrnovo.Value:=0;
      perc:=TrocaVirgPPto(floattostr(xprnovo.Value));
      if (CDSQuadraLote.Active) and (CDSLoteamento.Active=false) Then
      Begin
         CDSQuadraLote.First;
         if CDSQuadraLote.RecordCount>0 then
         begin
           while not CDSQuadraLote.Eof do begin
             varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
             CDSQuadraLote.Next;
           end;
         end;
         varquadra := copy(varquadra,1,length(varquadra)-1);

         ZQRecebimento.close;
         ZQRecebimento.SQL.clear;
         if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (not empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end

         else  if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end;
         ZQRecebimento.ExecSQL;
      end;

      if (CDSQuadraLote.Active) and (CDSLoteamento.Active) Then
      Begin
         CDSQuadraLote.First;
         if CDSQuadraLote.RecordCount>0 then
         begin
           while not CDSQuadraLote.Eof do begin
             varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
             CDSQuadraLote.Next;
           end;
         end;
         varquadra := copy(varquadra,1,length(varquadra)-1);


         CDSLoteamento.First;
         if CDSLoteamento.RecordCount>0 then
         begin
           while not CDSLoteamento.Eof do begin
             varloteamento := varloteamento +quotedstr(CDSLoteamentoidloteamento.Text)+',';
             CDSLoteamento.Next;
           end;
         end;
         varloteamento := copy(varloteamento,1,length(varloteamento)-1);


         ZQRecebimento.close;
         ZQRecebimento.SQL.clear;
         if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (not empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100) where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo=saldo+((saldo*'+perc+')/100),valor=valor+((valor*'+perc+')/100),proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end

         else  if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xprnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end;
         ZQRecebimento.ExecSQL;
      end;


    end
    else
    begin
      xprnovo.Value:=0;
      if (CDSQuadraLote.Active) and (CDSLoteamento.Active=false) Then
      Begin
         CDSQuadraLote.First;
         if CDSQuadraLote.RecordCount>0 then
         begin
           while not CDSQuadraLote.Eof do begin
             varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
             CDSQuadraLote.Next;
           end;
         end;
         varquadra := copy(varquadra,1,length(varquadra)-1);
         ZQRecebimento.close;
         ZQRecebimento.SQL.clear;
         if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (not empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end

         else  if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end;
         ZQRecebimento.ExecSQL;
      end;



      if (CDSQuadraLote.Active) and (CDSLoteamento.Active) Then
      Begin
         CDSQuadraLote.First;
         if CDSQuadraLote.RecordCount>0 then
         begin
           while not CDSQuadraLote.Eof do begin
             varquadra := varquadra +quotedstr(CDSQuadraLoteQuadra.Value)+',';
             CDSQuadraLote.Next;
           end;
         end;
         varquadra := copy(varquadra,1,length(varquadra)-1);

         CDSLoteamento.First;
         if CDSLoteamento.RecordCount>0 then
         begin
           while not CDSLoteamento.Eof do begin
             varloteamento := varloteamento +quotedstr(CDSLoteamentoidloteamento.Text)+',';
             CDSLoteamento.Next;
           end;
         end;
         varloteamento := copy(varloteamento,1,length(varloteamento)-1);

         ZQRecebimento.close;
         ZQRecebimento.SQL.clear;
         if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (not empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value>0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set saldo='+TrocaVirgPPto(floattostr(xvrnovo.Value))+', valor='+TrocaVirgPPto(floattostr(xvrnovo.Value))+',proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end

         else  if (not empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (dt_vencimento>= :dt) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
           ZQRecebimento.ParamByName('dt').AsDate:=strtodate(destadata.Text);
         end
         else if (empty(charrem(destadata.Text))) and (not empty(charrem(xreajuste.Text))) and (xvrnovo.Value=0) then
         begin
           if CBDesconsidera.Checked=false then
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
           else
              ZQRecebimento.SQL.Add(' update Recebimento set proximo_reajuste='+quotedstr(xreajuste.Text)+' where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
         end;
         ZQRecebimento.ExecSQL;
      end;

    end;
    ZQRecebimento.close;
    ZQRecebimento.SQL.clear;
    if (CDSQuadraLote.Active) and (CDSLoteamento.Active=false) Then
        ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento')
    else
        ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varloteamento+')'+') and (quadralote in ('+varquadra+')'+')  order by DT_Vencimento');
    ZQRecebimento.open;
  end;
  showmessage('Atualização completada !');
{      varquadra := '';
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
            ZQRecebimentosaldo.Value:=xvrnovo.Value;
            ZQRecebimento.post;
            ZQRecebimento.Next;
          end;
          ZQRecebimento.Refresh;
        //  ZQRecebtmp.next;
        //end;

        ZQRecebimento.close;
        ZQRecebimento.SQL.clear;
        if CBDesconsidera.Checked=false then
           ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'')  order by DT_Vencimento')
        else
           ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'') and (TipDoc=''BO'') and (saldo>''0'')  order by DT_Vencimento');

//        ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
        ZQRecebimento.Open;
        ZQRecebimento.First;
      end;
  end;}


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

procedure TFrm_Acerto_parcelas.dxButton13Click(Sender: TObject);
begin
  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;
  ZQRecebimento.First;
end;

procedure TFrm_Acerto_parcelas.dxButton14Click(Sender: TObject);
var
vlLinha,t: Integer;
begin
  Frm_Acerto_parcelas.barra2.Position:=0;

  with DBGreceb.DataSource.DataSet do
  begin
    First;
    t:=RecordCount - 1;
    Frm_Acerto_parcelas.barra2.Max:=t;
    for vlLinha := 0 to RecordCount - 1 do
    begin
     Frm_Acerto_parcelas.barra2.Position:=vlLinha;
     DBGreceb.SelectedRows.CurrentRowSelected := True;
     Next;
    end;
  end;
  DBGreceb.SelectedRows.Refresh;
  ZQRecebimento.First;
  Frm_Acerto_parcelas.barra2.Position:=0;
end;


procedure TFrm_Acerto_parcelas.DBGrid2DrawColumnCell(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.dxButton3Click(Sender: TObject);
var
ms:string;
begin
  Frm_Acerto_parcelas.barra2.Position:=0;
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


  ZQRecebtmp.first;
{  ZQRecebtmp.Filtered:=false;
  ZQRecebtmp.Filter:='Dt_Vencimento >= '+quotedstr('01/'+lmes.caption+'/'+xano.text)+ ' and Dt_Vencimento <= '+quotedstr(ms+'/'+lmes.caption+'/'+xano.text);
  ZQRecebtmp.Filtered:=true;
  ZQRecebtmp.first;}
    
  ZQRecebtmp2.first;
end;

procedure TFrm_Acerto_parcelas.EloteamentoExit(Sender: TObject);
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

procedure TFrm_Acerto_parcelas.DBGrid3KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSLoteamento.RecordCount>0) Then Begin
     CDSLoteamento.Delete;
  end;
end;

procedure TFrm_Acerto_parcelas.DBGrid2MouseUp(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.dxButton4Click(Sender: TObject);
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

procedure TFrm_Acerto_parcelas.dxButton5Click(Sender: TObject);
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

procedure TFrm_Acerto_parcelas.DBGrid5DrawColumnCell(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.DBGrid5MouseUp(Sender: TObject;
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

procedure TFrm_Acerto_parcelas.cbconferenciaClick(Sender: TObject);
begin
  if cbconferencia.Checked=true then
     DBGrid5.Visible:=true
  else
    DBGrid5.Visible:=false;   
end;


procedure TFrm_Acerto_parcelas.PrepararConsultaRecebimentoAntesDeAbrir(
  DataSet: TDataSet);
begin
  PrepararConsultaRecebimentoComNomes(DataSet);
end;

procedure TFrm_Acerto_parcelas.AfterConstruction;
begin
  inherited AfterConstruction;
  ZQRecebimento.BeforeOpen := PrepararConsultaRecebimentoAntesDeAbrir;
  ZQTempReceber.BeforeOpen := PrepararConsultaRecebimentoAntesDeAbrir;
  PrepararConsultaRecebimentoComNomes(ZQRecebimento);
  PrepararConsultaRecebimentoComNomes(ZQTempReceber);
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Acerto_parcelas, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Acerto_parcelas, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
