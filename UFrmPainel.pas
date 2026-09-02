unit UFrmPainel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, PReport, PRJpegImage, Vcl.ExtCtrls, PdfDoc,
  Vcl.StdCtrls,ShellAPI, Vcl.ComCtrls, Vcl.Mask, Vcl.CheckLst, XBanner;

type
  TfrmPainel = class(TForm)
    ZQCalculoEst: TZQuery;
    DS_CalculoEst: TDataSource;
    PReport1: TPReport;
    PRPage1: TPRPage;
    PRGridPanel1: TPRGridPanel;
    OpenDialog: TOpenDialog;
    PRText1: TPRText;
    Button3: TButton;
    Button1: TButton;
    RGFiltro: TRadioGroup;
    plperiodo: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    dtinicio: TDateTimePicker;
    dtfinal: TDateTimePicker;
    gbfinanc: TGroupBox;
    XBanner17: TXBanner;
    CLBEmpree: TCheckListBox;
    ZQLoteamento: TZQuery;


    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure RGFiltroClick(Sender: TObject);
    function  ListaEmpree: TStrings;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  frmPainel: TfrmPainel;

implementation

uses Tabelas, Principal, uRuntimeFields;

{$R *.dfm}

procedure TfrmPainel.Button1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmPainel.Button3Click(Sender: TObject);
var
x,varcontador,y,custo_compra,custovnd,abertorec,abertopg:integer;
totalparc,aberto,atrasado,baixado,paberto,patrasado,pbaixado,saldo,perc:double;
VarEmpree:String;
begin
  varcontador:=0;
  for x := 0 to CLBEmpree.Count-1 do begin
    if CLBEmpree.Checked[x] Then Begin
      inc(varcontador);
    end;
  end;
  if varcontador=0 then
  begin
    showmessage('Seleciona um empreendimento !');
    exit;
  end;

  if varcontador>1 then
  begin
    showmessage('Mais de um empreendimento selecionado. Deixe apenas um ! ');
    exit;
  end;
   Button3.Enabled:=false;
   PRPage1.Visible:=true;
   application.ProcessMessages;
   totalparc:=0;
   custo_compra:=0;
   custovnd:=0;
   abertorec:=0;
   abertopg:=0;
   saldo:=0;
   perc:=0;
   varcontador:=0;
   if FileExists(ExtractFilePath(Application.ExeName)+'painel.pdf') then
      deletefile(ExtractFilePath(Application.ExeName)+'painel.pdf');

  { ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('optimize table prod_est_quan');
   ZQCalculoEst.ExecSQL;}


  VarEmpree:= '';
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

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS Estatistica_imoveis');
   ZQCalculoEst.ExecSQL;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS Estatistica_imoveis AS (SELECT ');
   ZQCalculoEst.SQL.Add(' i.idimovel,i.loteamento_idloteamento,l.idloteamento,l.apelido, ');
   ZQCalculoEst.SQL.Add('       count(i.loteamento_idloteamento) as totallotes, ');
   ZQCalculoEst.SQL.Add('       sum(i.disponivel=''SIM'') as totdisponivel,sum(i.disponivel<>''SIM'') as totNaodisponivel ');
   ZQCalculoEst.SQL.Add(' FROM imovel i inner join loteamento l on l.idloteamento=i.loteamento_idloteamento where 1=1 ');
   if VarEmpree<>Emptystr then
      ZQCalculoEst.SQL.Add(' and i.loteamento_idloteamento in ('+VarEmpree+')');
   ZQCalculoEst.SQL.Add(' group by i.loteamento_idloteamento)');
 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
//   ZQCalculoEst.SQL.Add('select format(sum(vlr),2,''de_DE'') as total from Custo_estoque');
   ZQCalculoEst.SQL.Add('select * from Estatistica_imoveis');
   ZQCalculoEst.open;

   // Inserimos as informações de cabeçalho de
   // cada coluna
   ZQCalculoEst.First;
   // Varremos o cds e inserimos os dados
   // no PRText
   custo_compra:=ZQCalculoEst.FieldByName('totallotes').AsInteger;
   custovnd:=ZQCalculoEst.FieldByName('totdisponivel').AsInteger;
   abertorec:=ZQCalculoEst.FieldByName('totnaodisponivel').AsInteger;

   if (custovnd<>0) and (abertorec<>0) then
      perc:= (custovnd*100)/custo_compra
   else
     perc:=0;

   if (custovnd<>0) and (abertorec<>0) then
      saldo:= (abertorec*100)/custo_compra
   else
     saldo:=0;

   PRText1.Lines.Clear;
   if RGFiltro.ItemIndex=0 then
      PRText1.Lines.Add('Filtro Geral')
   else
     PRText1.Lines.Add('Filtro por Período: De: '+datetostr(dtinicio.Date)+'  até  '+datetostr(dtfinal.Date));

   PRText1.Lines.Add('');
   PRText1.Lines.Add('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Painel de Controle de Empresa <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<');
   PRText1.Lines.Add('');
   PRText1.Lines.Add('>>>>>>>>>>>>>>>>>>>>>>>>>>>> '+ZQCalculoEst.FieldByName('apelido').AsString+' <<<<<<<<<<<<<<<<<<<<<<<<<<<<');
   PRText1.Lines.Add('');
   PRText1.Lines.Add('Total de Lotes: '+FormatFloat('###,###,##0.##',custo_compra)+'  Lotes Disponíveis: '+FormatFloat('###,###,##0.##',custovnd)+'  Lotes não Disponíveis: '+FormatFloat('###,###,##0.##',abertorec));
   PRText1.Lines.Add('Disponíveis: '+FormatFloat('0.00',perc)+' %  - Não Disponíveis:'+FormatFloat('0.00',saldo)+' %');
   PRText1.Lines.Add('');
   PRText1.Lines.Add('');



   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS Estatistica_imoveis');
   ZQCalculoEst.ExecSQL;


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS Estatistica_imoveis AS (SELECT ');
   ZQCalculoEst.SQL.Add(' l.idloteamento,l.apelido,count(r.idloteamento) as totalParcelas, sum(r.saldo>0) as abertos, ');
   ZQCalculoEst.SQL.Add(' sum(r.saldo>0 and dt_vencimento<now()) as atrasados, sum(r.saldo=0) as baixados, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo>0)*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percaberto, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo>0 and dt_vencimento<now())*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percatrasado, ');
   ZQCalculoEst.SQL.Add(' concat(FORMAT(((sum(r.saldo=0)*100)/count(r.idloteamento)),2,''de_DE''),'' %'') as percbaixa ');
   ZQCalculoEst.SQL.Add('FROM recebimento r inner join loteamento l on l.idloteamento=r.idloteamento ');
   ZQCalculoEst.SQL.Add('  where dt_vencimento between :dt1 and :dt2');
   if VarEmpree<>Emptystr then
      ZQCalculoEst.SQL.Add(' and r.idloteamento in ('+VarEmpree+')');
   ZQCalculoEst.SQL.Add(' group by r.idloteamento order by l.apelido)');
   ZQCalculoEst.ParamByName('dt1').AsDate :=dtinicio.Date;
   ZQCalculoEst.ParamByName('dt2').AsDate :=dtfinal.Date;

 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('select * from Estatistica_imoveis');
   ZQCalculoEst.open;


   PRText1.Lines.Add('   Total de Parcelas : '+FormatFloat('###,###,##0.##',ZQCalculoEst.FieldByName('totalparcelas').AsInteger));
   PRText1.Lines.Add('Parcelas em Aberto: '+ZQCalculoEst.FieldByName('abertos').AsString+' - '+ZQCalculoEst.FieldByName('percaberto').AsString);
   PRText1.Lines.Add('Parcelas em Atraso: '+ZQCalculoEst.FieldByName('atrasados').AsString+' - '+ZQCalculoEst.FieldByName('percatrasado').AsString);
   PRText1.Lines.Add(' Parcelas Baixadas : '+ZQCalculoEst.FieldByName('baixados').AsString+' - '+ZQCalculoEst.FieldByName('percbaixa').AsString);
   PRText1.Lines.Add('');
   PRText1.Lines.Add('');


   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('DROP TEMPORARY TABLE IF EXISTS Estatistica_imoveis');
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS Estatistica_imoveis AS (SELECT ');
   ZQCalculoEst.SQL.Add('       l.idloteamento,l.apelido,round(sum(r.valor),2) as totalParcelas, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo > 0) THEN r.valor else 0 end),2) as aberto, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo > 0) and (r.dt_vencimento<now()) THEN r.valor else 0 end),2) as atrasado, ');
   ZQCalculoEst.SQL.Add('       round(sum(CASE WHEN (r.saldo = 0) THEN r.valor else 0 end),2) as pago ');
   ZQCalculoEst.SQL.Add(' FROM recebimento r inner join loteamento l on l.idloteamento=r.idloteamento ');
   ZQCalculoEst.SQL.Add('  where dt_vencimento between :dt1 and :dt2');
   if VarEmpree<>Emptystr then
      ZQCalculoEst.SQL.Add(' and r.idloteamento in ('+VarEmpree+')');
   ZQCalculoEst.SQL.Add(' group by r.idloteamento order by l.apelido)');
   ZQCalculoEst.ParamByName('dt1').AsDate :=dtinicio.Date;
   ZQCalculoEst.ParamByName('dt2').AsDate :=dtfinal.Date;

 //  showmessage(ZQCalculoEst.SQL.Text);
   ZQCalculoEst.ExecSQL;

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;
   ZQCalculoEst.SQL.Add('select * from Estatistica_imoveis');
   ZQCalculoEst.open;
   ZQCalculoESt.First;
   custo_compra:=0;
   aberto:=0;
   atrasado:=0;
   baixado:=0;
   paberto:=0;
   patrasado:=0;
   pbaixado:=0;
   totalparc:=0;
   totalparc:=ZQCalculoEst.FieldByName('totalParcelas').AsFloat;
   aberto:=ZQCalculoEst.FieldByName('aberto').AsFloat;
   atrasado:=ZQCalculoEst.FieldByName('atrasado').AsFloat;
   baixado:=ZQCalculoEst.FieldByName('pago').AsFloat;


   if (aberto<>0) then
      paberto:= (aberto*100)/totalparc
   else
     paberto:=0;

   if (atrasado<>0) then
      patrasado:= (atrasado*100)/totalparc
   else
     patrasado:=0;

   if (baixado<>0) then
      pbaixado:= (baixado*100)/totalparc
   else
     pbaixado:=0;


   PRText1.Lines.Add('  Total das Parcelas: R$ '+FormatFloat('###,###,##0.00',totalparc));
   PRText1.Lines.Add('Parcelas em Aberto: R$ '+FormatFloat('###,###,##0.00',aberto)+' - '+FormatFloat('###,###,##0.00',paberto)+' %');
   PRText1.Lines.Add('Parcelas em Atraso: R$ '+FormatFloat('###,###,##0.00',atrasado)+' - '+FormatFloat('###,###,##0.00',patrasado)+' %');
   PRText1.Lines.Add(' Parcelas Baixadas : R$ '+FormatFloat('###,###,##0.00',baixado)+' - '+FormatFloat('###,###,##0.00',pbaixado)+' %');

   ZQCalculoEst.Close;
   ZQCalculoEst.SQL.Clear;

   // Geramos e visualizamos o PDF

   PReport1.FileName:=ExtractFilePath(Application.ExeName)+'painel.pdf';
   PReport1.BeginDoc;
   PReport1.Print(PRPage1);
   PReport1.EndDoc;
   ShellExecute(Handle,'open',PChar(ExtractFilePath(Application.ExeName)+'painel.pdf'),nil,nil,SW_SHOWNORMAL);
   ZQCalculoEst.Close;
   PRPage1.Visible:=false;
   Button3.Enabled:=true;
   custo_compra:=0;
   custovnd:=0;
   abertorec:=0;
   abertopg:=0;
   saldo:=0;
   perc:=0;
end;
function TfrmPainel.ListaEmpree: TStrings;
begin
  Result := TStringList.Create;
  Result.Clear;
  Result.BeginUpdate;
  ZQLoteamento.open;
  ZQLoteamento.First;
  ZQLoteamento.DisableControls;
  while not ZQLoteamento.Eof do begin
        Result.Add(ZQLoteamento.FieldByName('apelido').AsString+' - '+ZQLoteamento.FieldByName('idloteamento').Text);
        ZQLoteamento.Next;
  end;
  ZQLoteamento.EnableControls;
  Result.EndUpdate;
  ZQLoteamento.close;
end;
procedure TfrmPainel.FormShow(Sender: TObject);
begin
  CLBEmpree.Clear;
  CLBEmpree.Items:= ListaEmpree;
  dtinicio.date:=date;
  dtfinal.date:=date;
end;

procedure TfrmPainel.RGFiltroClick(Sender: TObject);
begin
  if RGFiltro.ItemIndex=0 then
     plperiodo.Visible:=false
  else
     plperiodo.Visible:=true;
end;


procedure TfrmPainel.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TfrmPainel, 'ZQLoteamento', False);
  RegisterRuntimeField(TfrmPainel, 'ZQLoteamento', 'ZQLoteamentoidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TfrmPainel, 'ZQLoteamento', 'ZQLoteamentoapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
