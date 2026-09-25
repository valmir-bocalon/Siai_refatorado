unit ReajusteDeParcelasold;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, XNum, XLabel3D, Mask, XDate, StdCtrls, Grids, DBGrids,
    dxButton, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBCtrls, TFlatGaugeUnit,DbiProcs,Shellapi;

type
  TFrm_ReajusteDeParcelasold = class(TForm)
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
    barra: TFlatGauge;
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
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_ReajusteDeParcelasold: TFrm_ReajusteDeParcelasold;

implementation

uses tabelas, PesqRecebimento,funcoes, AchaIgpm, PesqRecebimento2, uRuntimeFields;

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
      Frm_ReajusteDeParcelas.barra.Progress:=0;
      t:=BookMarkList.Count -1;
      Frm_ReajusteDeParcelas.barra.MaxValue:=t;

      for i := 0 to BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Frm_ReajusteDeParcelas.barra.Progress:=i;
        Strings.Add(DataSet.FindField(FieldName).AsString);
        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=false;
        Frm_ReajusteDeParcelas.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=true;
        Frm_ReajusteDeParcelas.ZQRecebimento.First;
        if Frm_ReajusteDeParcelas.ZQRecebimento.recordcount>0 then
        begin
//          Frm_ReajusteDeParcelas.barra.MaxValue:=Frm_ReajusteDeParcelas.ZQRecebimento.recordcount;
          while not Frm_ReajusteDeParcelas.ZQRecebimento.Eof do
          begin
//            Frm_ReajusteDeParcelas.barra.Progress:=Frm_ReajusteDeParcelas.ZQRecebimento.recno;
            Frm_ReajusteDeParcelas.ZQRecebtmp.Insert;
            Frm_ReajusteDeParcelas.ZQRecebtmpidrecebimento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('idrecebimento').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpdocumento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('documento').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpcliente.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('cliente').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpusuario.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('usuario').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpDt_Entrada.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Dt_Entrada').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpDt_Vencimento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpValor.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Valor').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpVrDoc.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('VrDoc').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpordem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('ordem').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpTipDoc.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('TipDoc').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpsaldo.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('saldo').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmporigem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('origem').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpadversa.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('adversa').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmprecpag.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('recpag').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpnumordem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('numordem').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpadversanome.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('adversanome').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpidloteamento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('idloteamento').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpvenda_idvenda.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('venda_idvenda').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpquadralote.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('quadralote').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpnomeadversa.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('nomeadversa').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmpParcelas_fixas.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Parcelas_fixas').Value;
            Frm_ReajusteDeParcelas.ZQRecebtmp.Post;
            Frm_ReajusteDeParcelas.ZQRecebimento.Next;
          end;
        end;
        Frm_ReajusteDeParcelas.barra.Progress:=0;
        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=false;
      end;
    finally
      GotoBookMark(bm);
      FreeBookMark(bm);
      Frm_ReajusteDeParcelas.barra.Progress:=0;
      EnableControls;
    end;
  end; // With DataSet do
end;


procedure TFrm_ReajusteDeParcelasold.DXBFecharClick(Sender: TObject);
begin
  close;
end;
procedure TFrm_ReajusteDeParcelasold.EadversaExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelasold.EloteExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelasold.FormShow(Sender: TObject);
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


end;
procedure TFrm_ReajusteDeParcelasold.botoes_setas;
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

procedure TFrm_ReajusteDeParcelasold.DXBPrimeiroClick(Sender: TObject);
begin
  ZQRecebimento.First;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelasold.DXBAnteriorClick(Sender: TObject);
begin
  ZQRecebimento.Prior;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelasold.DXBProximoClick(Sender: TObject);
begin
  ZQRecebimento.Next;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelasold.DXBUltimoClick(Sender: TObject);
begin
  ZQRecebimento.Last;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelasold.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento2.Top := Frm_ReajusteDeParcelas.Top+100;
  FrmPesqRecebimento2.Left := Frm_ReajusteDeParcelas.Left;
  FrmPesqRecebimento2.showmodal;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelasold.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrm_ReajusteDeParcelasold.FormClose(Sender: TObject;
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

procedure TFrm_ReajusteDeParcelasold.DBGrid4KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelasold.DBGrid1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelasold.dxButton2Click(Sender: TObject);
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
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
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
         ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
      else
         ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
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
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      mensagem('Sem Filtro Definido.');
  percent.SetFocus;
end;

procedure TFrm_ReajusteDeParcelasold.DBGRecebDrawColumnCell(Sender: TObject;
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

end;

procedure TFrm_ReajusteDeParcelasold.dxButton1Click(Sender: TObject);
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
  if (percent.Value>0) and (ZQRecebtmp.RecordCount>0) then
  begin
    dxButton1.Enabled:=false;
    ZQRecebtmp.First;
    barra.Progress:=0;
    barra.MaxValue:=ZQRecebtmp.RecordCount;
    while not ZQRecebtmp.Eof do
    begin
      barra.progress:=ZQRecebtmp.RecNo;
      vr:=ZQRecebtmpvalor.Value;
      sld:=ZQRecebtmpsaldo.Value;
      corrigido:=vr+((vr*percent.Value)/100);
      corrigido2:=sld+((sld*percent.Value)/100);
      ZQRecebtmp.Edit;
      ZQRecebtmpsaldo_corr.Value:=ExRound(corrigido,2);
      ZQRecebtmp.post;
      ZQRecebtmp.Next;
    end;
    if simnao('Confirma o Reajuste das Parcelas ?','SIM') then
    begin
      vr:=0;
      sld:=0;
      corrigido:=0;
      corrigido2:=0;

      varquadra := '';
      ZQRecebtmp.First;
      if ZQRecebtmp.RecordCount>0 then
      begin
        while not ZQRecebtmp.Eof do
        begin
          varquadra := varquadra +quotedstr(ZQRecebtmpquadralote.Value)+',';
          ZQRecebtmp.Next;
        end;
        ZQRecebtmp.First;
        ZQRecebtmp.IndexName:='vencimento';
        ZQRecebtmp.First;
        varquadra := copy(varquadra,1,length(varquadra)-1);
        DBGReceb.SelectedRows.Clear;
//        while not ZQRecebtmp.eof do
//        begin
          ZQRecebimento.close;
          ZQRecebimento.SQL.clear;
          if CBDesconsidera.Checked=false then
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento')
          else
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento');

//          ZQRecebimento.SQL.Add(' Select * from Recebimento where Parcelas_fixas<>''S'' and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt order by DT_Vencimento');
          ZQRecebimento.ParamByName('dt').AsDate:=ZQRecebtmpDt_Vencimento.Value;
          ZQRecebimento.Open;
          ZQRecebimento.First;
          barra.Progress:=0;
          barra.MaxValue:=ZQRecebimento.RecordCount;
          ZQRecebimento.DisableControls;
          while not ZQRecebimento.Eof do
          begin
            barra.progress:=ZQRecebimento.RecNo;
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


  ZQRecebtmp.First;
  ql:=ZQRecebtmpquadralote.Value;
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
  barra.Progress:=0;
  barra.MaxValue:=ZQRecebtmp.RecordCount;
  while not ZQRecebtmp.Eof do
  begin
    barra.progress:=ZQRecebtmp.RecNo;
    if ql=ZQRecebtmpquadralote.Value then
    begin
      ZQRecebtmp.Next;
    end
    else
    begin
      ql:=ZQRecebtmpquadralote.Value;
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
  ql:=ZQRecebtmpquadralote.Value;
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
  barra.Progress:=0;


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
  barra.Progress:=0;
end;

procedure TFrm_ReajusteDeParcelasold.xmesExit(Sender: TObject);
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
procedure TFrm_ReajusteDeParcelasold.dxButton13Click(Sender: TObject);
begin
  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;
  ZQRecebimento.First;
end;

procedure TFrm_ReajusteDeParcelasold.dxButton14Click(Sender: TObject);
var
vlLinha,t: Integer;
begin
  Frm_ReajusteDeParcelas.barra.Progress:=0;
  with DBGreceb.DataSource.DataSet do
  begin
    First;
    t:=RecordCount - 1;
    Frm_ReajusteDeParcelas.barra.MaxValue:=t;
    for vlLinha := 0 to RecordCount - 1 do
    begin
     Frm_ReajusteDeParcelas.barra.Progress:=vlLinha;
     DBGreceb.SelectedRows.CurrentRowSelected := True;
     Next;
    end;
  end;
  DBGreceb.SelectedRows.Refresh;
  ZQRecebimento.First;
  Frm_ReajusteDeParcelas.barra.Progress:=0;
end;

procedure TFrm_ReajusteDeParcelasold.DBGrid2DrawColumnCell(Sender: TObject;
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

end;

procedure TFrm_ReajusteDeParcelasold.dxButton3Click(Sender: TObject);
begin
  Frm_ReajusteDeParcelas.barra.Progress:=0;
  ZQRecebtmp.close;
  ZQRecebtmp.createdataset;
  GetSelectedRecord(ZQRecebimento, 'idrecebimento',DBGReceb.SelectedRows, ListBox1.Items);
  dxButton1.Enabled:=true;
end;

procedure TFrm_ReajusteDeParcelasold.EloteamentoExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelasold.GroupBox5Enter(Sender: TObject);
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

//  mes_anos:=MesAno(date);
  Frm_AchaIgpm.ECidade.Text := mes_anos;
  DM_tabelas.ZQAchaIgpm.open;
  Frm_AchaIgpm.ShowModal;
  percent.Value:=DM_tabelas.ZQAchaIgpm.FieldByName('percentual').Value;
  DM_tabelas.ZQAchaIgpm.close;
end;

procedure TFrm_ReajusteDeParcelasold.percentExit(Sender: TObject);
begin
  dxbutton3.Enabled:=true;
  dxButton13.Enabled:=true;
  dxButton14.Enabled:=true;
  dxButton14.SetFocus;
end;

procedure TFrm_ReajusteDeParcelasold.DBGrid3KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSLoteamento.RecordCount>0) Then Begin
     CDSLoteamento.Delete;
  end;
end;


procedure TFrm_ReajusteDeParcelasold.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1quadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1lote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1lado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradasomar', 'somar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParceladocumento', 'documento', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaordem', 'ordem', TStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaorigem', 'origem', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TStringField, fkData, 11, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelasold, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreelote', 'lote', TStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreelado', 'lado', TStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreematri', 'matri', TStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelasold, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
