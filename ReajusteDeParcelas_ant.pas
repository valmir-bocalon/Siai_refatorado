unit ReajusteDeParcelas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XBanner, XNum, XLabel3D, Mask, XDate, StdCtrls, Grids, DBGrids,
    dxButton, DB, DBClient, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, DBCtrls, TFlatGaugeUnit,DbiProcs,Shellapi, ImgList, ExtCtrls,
  dxCore2;

type
  TFrm_ReajusteDeParcelas = class(TForm)
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
    lbreg: TLabel;
    lblreg: TLabel;
    Panel1: TPanel;




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
    procedure DBGRecebColEnter(Sender: TObject);
    procedure DBGRecebMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_ReajusteDeParcelas: TFrm_ReajusteDeParcelas;

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
      i := 0;
//      for i := 0 to BookMarkList.Count -1 do
      while i <= BookMarkList.Count -1 do
      begin
        BookMark := BookMarkList[i];
        Frm_ReajusteDeParcelas.barra.Progress:=i;
        Strings.Add(DataSet.FindField(FieldName).AsString);
{        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=false;
        Frm_ReajusteDeParcelas.ZQRecebimento.Filter:='idrecebimento='+quotedstr(DataSet.FindField(FieldName).AsString);
        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=true;
        Frm_ReajusteDeParcelas.ZQRecebimento.First;
        if Frm_ReajusteDeParcelas.ZQRecebimento.recordcount>0 then}
        Frm_ReajusteDeParcelas.ZQRecebimento.Locate('idrecebimento',DataSet.FindField(FieldName).AsString,[]);
        if Frm_ReajusteDeParcelas.ZQRecebimento.Found then
        begin
//          Frm_ReajusteDeParcelas.barra.MaxValue:=Frm_ReajusteDeParcelas.ZQRecebimento.recordcount;
//          while not Frm_ReajusteDeParcelas.ZQRecebimento.Eof do
  //        begin
//            Frm_ReajusteDeParcelas.barra.Progress:=Frm_ReajusteDeParcelas.ZQRecebimento.recno;
            Frm_ReajusteDeParcelas.ZQRecebtmp.Append;
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
            //Frm_ReajusteDeParcelas.ZQRecebimento.Next;
          //end;
        end;
//        Frm_ReajusteDeParcelas.ZQRecebimento.Filtered:=false;
        inc(i);
      end;
    finally
      Frm_ReajusteDeParcelas.barra.Progress:=0;
      GotoBookMark(bm);
      FreeBookMark(bm);
      Frm_ReajusteDeParcelas.barra.Progress:=0;
      EnableControls;
    end;
  end; // With DataSet do
end;


procedure TFrm_ReajusteDeParcelas.DXBFecharClick(Sender: TObject);
begin
  close;
end;
procedure TFrm_ReajusteDeParcelas.EadversaExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelas.EloteExit(Sender: TObject);
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
      Showmessage('Esta Quadra e lote n�o foram vendidos!!!!');
    Equadra.Text := '';
    Elote.text := '';
    Equadra.SetFocus;
  end;
end;

procedure TFrm_ReajusteDeParcelas.FormShow(Sender: TObject);
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
procedure TFrm_ReajusteDeParcelas.botoes_setas;
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

procedure TFrm_ReajusteDeParcelas.DXBPrimeiroClick(Sender: TObject);
begin
  ZQRecebimento.First;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas.DXBAnteriorClick(Sender: TObject);
begin
  ZQRecebimento.Prior;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas.DXBProximoClick(Sender: TObject);
begin
  ZQRecebimento.Next;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas.DXBUltimoClick(Sender: TObject);
begin
  ZQRecebimento.Last;
  botoes_setas;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas.DXBPesquisarClick(Sender: TObject);
begin
//  FrmAchaReceb.showmodal;
  FrmPesqRecebimento2.Top := Frm_ReajusteDeParcelas.Top+100;
  FrmPesqRecebimento2.Left := Frm_ReajusteDeParcelas.Left;
  FrmPesqRecebimento2.showmodal;
  DBGReceb.SetFocus;

end;

procedure TFrm_ReajusteDeParcelas.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrm_ReajusteDeParcelas.FormClose(Sender: TObject;
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

procedure TFrm_ReajusteDeParcelas.DBGrid4KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSAdversaTemp2.RecordCount>0) Then Begin
    CDSAdversaTemp2.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelas.DBGrid1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSQuadraLote.RecordCount>0) Then Begin
    CDSQuadraLote.Delete;
  end;

end;

procedure TFrm_ReajusteDeParcelas.dxButton2Click(Sender: TObject);
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
  if mmes='Mar�o' then
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
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text))
       else
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text));
       ZQRecebimento.ExecSQL;


       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
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
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+')')
       else
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+')');
       ZQRecebimento.ExecSQL;



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
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+')')
       else
          ZQRecebimento.SQL.Add(' update Recebimento set marcar=''0'' where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+')');
       ZQRecebimento.ExecSQL;


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
  lbreg.Caption:=inttostr(ZQRecebimento.RecordCount);
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      showmessage('Sem Filtro Definido.');
  percent.SetFocus;
end;

procedure TFrm_ReajusteDeParcelas.DBGRecebDrawColumnCell(Sender: TObject;
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



  if ZQRecebimento.active=true then
  begin
    if DBEIDReceb.Text = ZQRecebimento.FieldByName('idrecebimento').Text Then Begin
      DBGReceb.Canvas.Brush.Color := $006CFFFF;
      DBGReceb.Canvas.Font.Color := $00A80000;
    end;
    DBGReceb.DefaultDrawDataCell(Rect, DBGReceb.columns[datacol].field, State);
    if Column.Field = ZQRecebimento.FieldByName('marcar') then
    begin
      DBGReceb.Canvas.FillRect(Rect);
      if ZQRecebimento.FieldByName('marcar').Value = '0' then
        ImageList1.Draw(DBGReceb.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
      else
        ImageList1.Draw(DBGReceb.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
    end;
  end;
end;

procedure TFrm_ReajusteDeParcelas.dxButton1Click(Sender: TObject);
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
      Panel1.Visible:=true;
      application.ProcessMessages;
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
{          ZQRecebimento.close;
          ZQRecebimento.SQL.clear;
          if CBDesconsidera.Checked=false then
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento')
          else
             ZQRecebimento.SQL.Add(' Select * from Recebimento where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento');

          ZQRecebimento.ParamByName('dt').AsDate:=ZQRecebtmpDt_Vencimento.Value;
          ZQRecebimento.Open;
          ZQRecebimento.First;
          barra.Progress:=0;
          barra.MaxValue:=ZQRecebimento.RecordCount;
          ZQRecebimento.DisableControls;
          while not ZQRecebimento.Eof do
          begin
            barra.progress:=ZQRecebimento.RecNo;
            vr:=ZQRecebimentovalor.Value;
            sld:=ZQRecebimentosaldo.Value;
            corrigido:=vr+((vr*percent.Value)/100);
            corrigido2:=sld+((sld*percent.Value)/100);
            ZQRecebimento.Edit;
            ZQRecebimentovalor.Value:=ExRound(corrigido,2);
            ZQRecebimentosaldo.Value:=ExRound(corrigido2,2);
            ZQRecebimentoReajustado.Value:='SIM';
            ZQRecebimentoData_reajuste.Value:=date;
            ZQRecebimentoProximo_Reajuste.Value:=copy(datetostr(data),4,7);
            ZQRecebimento.post;
            ZQRecebimento.Next;
          end;}


          ZQRecebimento.close;
          ZQRecebimento.SQL.clear;
          if CBDesconsidera.Checked=false then
          begin
             ZQRecebimento.SQL.Add(' update Recebimento set marcar=''1'',Percentual_reajuste='+StringReplace(floattostr(percent.Value), ',', '.', [])+',valor=round(valor+(valor*'+StringReplace(floattostr(percent.Value), ',', '.', [])+'/100),2),sld_antes_reajuste=saldo,saldo=round(saldo+(saldo*'+StringReplace(floattostr(percent.Value), ',', '.', [])+'/100),2),reajustado=''SIM'',data_reajuste= :dr,Proximo_Reajuste='+quotedstr(copy(datetostr(data),4,7))+' where ((numboleto is null) or (numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento');
          end
          else
          begin
             ZQRecebimento.SQL.Add(' update Recebimento set marcar=''1'',Percentual_reajuste='+StringReplace(floattostr(percent.Value), ',', '.', [])+',valor=round(valor+(valor*'+StringReplace(floattostr(percent.Value), ',', '.', [])+'/100),2),sld_antes_reajuste=saldo,saldo=round(saldo+(saldo*'+StringReplace(floattostr(percent.Value), ',', '.', [])+'/100),2),reajustado=''SIM'',data_reajuste= :dr,Proximo_Reajuste='+quotedstr(copy(datetostr(data),4,7))+' where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (quadralote in ('+varquadra+')'+') and (DT_Vencimento >= :dt) order by DT_Vencimento');
          end;
          ZQRecebimento.ParamByName('dt').AsDate:=ZQRecebtmpDt_Vencimento.Value;
          ZQRecebimento.ParamByName('dr').AsDate:=date;
          ZQRecebimento.ExecSQL;

//          ZQRecebimento.Refresh;
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
    showmessage('Informe o Percentual; Adicione os T�tulos Selecionados no Grid !');
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
  if DM_Tabelas.ZQVenda.Active=true then
     DM_Tabelas.ZQVenda.Close;

  ZQEmpree.close;
  ZQVenda.close;
  ZQEntrada.close;
  ZQParcela.close;
  ql:='';
  // at� aqui
  if DM_tabelas.ZQRecebimento.Active=true then
     DM_tabelas.ZQRecebimento.refresh;

  dxButton13.Enabled:=false;
  dxButton14.Enabled:=false;
  dxButton3.Enabled:=false;
  dxButton1.Enabled:=false;
  ZQRecebimento.EnableControls;
  barra.Progress:=0;
  Panel1.Visible:=false;
  application.ProcessMessages;

  Showmessage('Reajuste Terminado Com Sucesso!');
end;

procedure TFrm_ReajusteDeParcelas.xmesExit(Sender: TObject);
var
mmes,ms,xdata:string;
ds:Tdatetime;

begin
  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Mar�o' then
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
procedure TFrm_ReajusteDeParcelas.dxButton13Click(Sender: TObject);
var
vlLinha,t: Integer;
varquadra,vardevedor,varlotea:string;
mmes,ms,xdata:string;
ds:Tdatetime;


begin
{  DBGReceb.SelectedRows.Clear;
  ListBox1.Clear;
  ZQRecebimento.First;}

  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Mar�o' then
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
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento' );
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where  (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where  (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where  (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
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
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento' );
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where  ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


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
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


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
  lbreg.Caption:=inttostr(ZQRecebimento.RecordCount);
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      showmessage('Sem Filtro Definido.')
  else
     dxButton3.Enabled:=false;
end;


procedure TFrm_ReajusteDeParcelas.dxButton14Click(Sender: TObject);
var
vlLinha,t: Integer;
varquadra,vardevedor,varlotea:string;
mmes,ms,xdata:string;
ds:Tdatetime;

begin
{  Frm_ReajusteDeParcelas.barra.Progress:=0;
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
  Frm_ReajusteDeParcelas.barra.Progress:=0;}

  mmes:=alltrim(xmes.Text);
  if mmes='Janeiro' then
     xdata:='01';
  if mmes='Fevereiro' then
     xdata:='02';
  if mmes='Mar�o' then
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
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}

       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento' );
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where  (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and (numboleto is null or numboleto='''') and (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and  (Parcelas_fixas<>''S'' or Parcelas_fixas is null ) and  TipDoc=''BO'' and saldo>''0'' and quadralote in ('+varquadra+') and Proximo_Reajuste ='+quotedstr(lmes.caption+'/'+xano.text)+' order by DT_Vencimento');
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
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento' );
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where  ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


      ZQRecebimento.close;
      ZQRecebimento.SQL.clear;
{      ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and adversa in ('+vardevedor+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
      ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
      ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
      if CBDesconsidera.Checked=false then
         ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
      else
         ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (adversa in ('+vardevedor+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
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
       if CBDesconsidera.Checked=false then
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ');
         ZQRecebimento.SQL.Add(' where  ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end
       else
       begin
         ZQRecebimento.SQL.Add('update recebimento set marcar= case when (marcar=''1'') then ''0'' when (marcar=''0'') then ''1'' end ' );
         ZQRecebimento.SQL.Add(' where ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       end;
       ZQRecebimento.ExecSQL;


       ZQRecebimento.close;
       ZQRecebimento.SQL.clear;
{       ZQRecebimento.SQL.Add(' Select * from Recebimento where TipDoc=''BO'' and saldo>''0'' and idloteamento in ('+varlotea+')'+' and DT_Vencimento >= :dt and DT_Vencimento <= :dt2 order by DT_Vencimento');
       ZQRecebimento.ParamByName('dt').AsDate:=strtodate('01/'+lmes.caption+'/'+xano.text);
       ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(ms+'/'+lmes.caption+'/'+xano.text);}
       if CBDesconsidera.Checked=false then
          ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and ((numboleto is null or numboleto='''')) and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento')
       else
          ZQRecebimento.SQL.Add(' Select * from Recebimento where marcar=''0'' and ((Parcelas_fixas<>''S'' or Parcelas_fixas is null )) and (TipDoc=''BO'') and (saldo>''0'') and (idloteamento in ('+varlotea+')'+') and (Proximo_Reajuste='+quotedstr(lmes.caption+'/'+xano.text)+') order by DT_Vencimento');
       ZQRecebimento.Open;
       ZQRecebimento.First;
     end;
  end;
  lbreg.Caption:=inttostr(ZQRecebimento.RecordCount);  
  if (CDSQuadraLote.RecordCount=0) and (CDSAdversaTemp2.RecordCount=0) and (CDSLoteamento.RecordCount=0) Then
      showmessage('Sem Filtro Definido.')
  else
     dxButton3.Enabled:=true;
end;



procedure TFrm_ReajusteDeParcelas.DBGrid2DrawColumnCell(Sender: TObject;
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

procedure TFrm_ReajusteDeParcelas.dxButton3Click(Sender: TObject);
begin
{  Frm_ReajusteDeParcelas.barra.Progress:=0;
  ZQRecebtmp.close;
  ZQRecebtmp.createdataset;
  GetSelectedRecord(ZQRecebimento, 'idrecebimento',DBGReceb.SelectedRows, ListBox1.Items);}

  ZQRecebimento.First;
  barra.MaxValue:=ZQRecebimento.recordcount;
  while not ZQRecebimento.Eof do
  begin
    barra.Progress:=ZQRecebimento.recno;
    ZQRecebtmp.Append;
    ZQRecebtmpidrecebimento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('idrecebimento').Value;
    ZQRecebtmpdocumento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('documento').Value;
    ZQRecebtmpcliente.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('cliente').Value;
    ZQRecebtmpusuario.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('usuario').Value;
    ZQRecebtmpDt_Entrada.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Dt_Entrada').Value;
    ZQRecebtmpDt_Vencimento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value;
    ZQRecebtmpValor.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Valor').Value;
    ZQRecebtmpVrDoc.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('VrDoc').Value;
    ZQRecebtmpordem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('ordem').Value;
    ZQRecebtmpTipDoc.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('TipDoc').Value;
    ZQRecebtmpsaldo.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('saldo').Value;
    ZQRecebtmporigem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('origem').Value;
    ZQRecebtmpadversa.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('adversa').Value;
    ZQRecebtmprecpag.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('recpag').Value;
    ZQRecebtmpnumordem.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('numordem').Value;
    ZQRecebtmpadversanome.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('adversanome').Value;
    ZQRecebtmpidloteamento.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('idloteamento').Value;
    ZQRecebtmpvenda_idvenda.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('venda_idvenda').Value;
    ZQRecebtmpquadralote.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('quadralote').Value;
    ZQRecebtmpnomeadversa.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('nomeadversa').Value;
    ZQRecebtmpParcelas_fixas.Value:=Frm_ReajusteDeParcelas.ZQRecebimento.FieldByName('Parcelas_fixas').Value;
    ZQRecebtmp.Post;
    ZQRecebimento.Next;
  end;
  lblreg.Caption:=inttostr(ZQRecebtmp.RecordCount);
  dxButton1.Enabled:=true;
end;

procedure TFrm_ReajusteDeParcelas.EloteamentoExit(Sender: TObject);
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

procedure TFrm_ReajusteDeParcelas.GroupBox5Enter(Sender: TObject);
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
  if mmes='Mar�o' then
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

procedure TFrm_ReajusteDeParcelas.percentExit(Sender: TObject);
begin
  dxbutton3.Enabled:=true;
  dxButton13.Enabled:=true;
  dxButton14.Enabled:=true;
  dxButton14.SetFocus;
end;

procedure TFrm_ReajusteDeParcelas.DBGrid3KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (IntToStr(key)='46') and (CDSLoteamento.RecordCount>0) Then Begin
     CDSLoteamento.Delete;
  end;
end;

procedure TFrm_ReajusteDeParcelas.DBGRecebColEnter(Sender: TObject);
begin
   if DBGReceb.SelectedField = ZQRecebimento.FieldByName('marcar') then
    DBGReceb.Options := DBGReceb.Options - [dgEditing]
  else
    DBGReceb.Options := DBGReceb.Options + [dgEditing];

end;

procedure TFrm_ReajusteDeParcelas.DBGRecebMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBGReceb.SelectedField.fieldname = 'marcar' Then Begin
    ZQRecebimento.Edit;
    if  ZQRecebimento.FieldByName('marcar').Value = '0' Then
      ZQRecebimento.FieldByName('marcar').Value := '1'
    else
      ZQRecebimento.FieldByName('marcar').Value := '0';
//    Atualiza_tela;
  end;

end;


procedure TFrm_ReajusteDeParcelas.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemptipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemplocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatemprenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQAdvsersatemp', 'ZQAdvsersatempcadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQVen1', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1idvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1datavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1imovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1valorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1idimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1loteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1quadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1lote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1lado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1medidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1medidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1medidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1medidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1area', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1confrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1confrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1confrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1confrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1esquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1esquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1esquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1valorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1Obervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1cornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVen1', 'ZQVen1disponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQTempReceber', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebersaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebermarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebertipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberlocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberprofissao', 'profissao', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberplanomascara', 'planomascara', TWideStringField, fkLookup, 0, 0, False, '', '', '', '', 0, 'contabil', 'DM_Tabelas.ZQPlanoDeContas', 'codigo', 'mascara', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebercadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebervenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebernomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempRecebersld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQTempReceber', 'ZQTempReceberPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentosaldocalc', 'saldocalc', TFloatField, fkCalculated, 0, 0, False, '', '#,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentonomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '!99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentosld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQRecebimento', 'ZQRecebimentoPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQEntrada', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQParcela', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQVenda', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_ReajusteDeParcelas, 'ZQEmpree', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreequadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreelote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreelado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreemedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreemedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreemedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreemedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreearea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreevalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreeObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreecornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreedisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreematri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_ReajusteDeParcelas, 'ZQEmpree', 'ZQEmpreematricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
