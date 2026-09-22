unit UAgenda;

interface

uses
Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, XPMan, ComCtrls, ExtCtrls, StdCtrls, Menus, Mask, ActnMan,
ActnColorMaps, Grids, Buttons, DBCtrls,Dateutils, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client,
  FireDAC.Phys.MySQL, FireDAC.Phys.MySQLDef, AdvUtil, AdvObj, BaseGrid, AdvGrid,
  DBAdvGrid, AdvPDFIO, AdvGridPDFIO, Vcl.DBGrids;

type
  TFrmAgenda = class(TForm)
    ds_qryagenda: TDataSource;
    Panel1: TPanel;
    LabelComp: TLabel;
    Panel2: TPanel;
    StringGridCal: TStringGrid;
    Panel3: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    BitBtn2: TBitBtn;
    BitBtn1: TBitBtn;
    DBRichEdit1: TDBRichEdit;
    function CalcFeriadosMoveis(aYear: Integer): TDate;
    procedure FormCreate(Sender: TObject);
    procedure Dias(Mes: Integer;Ano: Integer);
    procedure BitBtn1Click(Sender: TObject);
    procedure StringGridCalDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure StringGridCalSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure DBAdvGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure StringGridCalDblClick(Sender: TObject);
    procedure ds_qryagendaStateChange(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmAgenda: TFrmAgenda;
  Meses: Integer = 1; //Variavel global para contar passagem pelos meses
  mes2,anos,data:string;
  xtexto:String;

implementation
uses tabelas, UAgenda_mov, uRuntimeFields, Funcoes;

{$R *.dfm}

function TFrmAgenda.CalcFeriadosMoveis(aYear: Integer): TDate;
var Month, Day,
   A, B, C, D, E: Integer;
begin
   A := aYear mod 4;
   B := aYear mod 7;
   C := aYear mod 19;
   D := (19 * C + 24) mod 30;
   E := (2 * A + 4 * B + 6 * D + 5) mod 7;
   if (22 + D + E) > 31 then
   begin
      Day := D + E - 9;
      Month := 4;
   end
   else
   begin
      Day := D + E + 22;
      Month := 3;
   end;
   Result := EncodeDate(aYear, Month, Day);
end;

function StrZero(iNumero, iComp: Integer): String;
begin
  Result := StringOfChar('0',iComp-Length(IntToStr(iNumero)))+IntToStr(iNumero);
end;
function MESEXT(num: string): string;
begin
  if num = '01' then
  MESEXT:='Janeiro';
  if num = '02' then
  MESEXT:='Fevereiro';
  if num = '03' then
  MESEXT:='Março';
  if num = '04' then
  MESEXT:='Abril';
  if num = '05' then
  MESEXT:='Maio';
  if num = '06' then
  MESEXT:='Junho';
  if num = '07' then
  MESEXT:='Julho';
  if num = '08' then
  MESEXT:='Agosto';
  if num = '09' then
  MESEXT:='Setembro';
  if num = '10' then
  MESEXT:='Outubro';
  if num = '11' then
  MESEXT:='Novembro';
  if num = '12' then
  MESEXT:='Dezembro';
end;
procedure TFrmAgenda.BitBtn1Click(Sender: TObject);
var
xano:integer;
begin
//   anos:=FormatDateTime('yyyy', Date );
   Inc(Meses); //Soma 1 a variavel meses, Cada click mostra o próximo mês.
   if Meses >= 13 then
   begin
      xano:=anos.ToInteger;
      Meses := 1;
      inc(xano);
      anos:=xano.ToString;
   end;

   Dias(Meses,anos.ToInteger); //Agora ainda terá de decidir como será definido o ano.//Poderá utilizar uma varaivel no lugar onde está escrito 2015 para passar o ano para//o procedimento que escreverá no grid.

end;

procedure TFrmAgenda.BitBtn2Click(Sender: TObject);
var
xano:integer;
begin

//   anos:=FormatDateTime('yyyy', Date );
   if Meses = 1 then
   begin
      xano:=anos.ToInteger;
      Meses := 12;
      dec(xano);
      anos:=xano.ToString;
   end;
   dec(Meses);
   Dias(Meses,anos.ToInteger); //Agora ainda terá de decidir como será definido o ano.//Poderá utilizar uma varaivel no lugar onde está escrito 2015 para passar o ano para//o procedimento que escreverá no grid.
end;

procedure TFrmAgenda.DBAdvGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
//          Brush.Color := $006A9BFF; // <--- essa linha coloca AZUL
//          AFont.Color := ClBlack;
  //        AFont.Style := [];
end;

procedure TFrmAgenda.Dias(Mes: Integer;Ano: Integer);
var
days: array[0..6] of string;
x, i, iNumDays, iDay: Integer; //Acrescentada variavel x para loop...
iRowCtr, iColCtr: Integer;
xtexto1:string;
iAno:TdateTime;
begin//Loops para limpar celulas do Grid antes de preencher o proximo mes, pode ser removido para ver o que ocorre.
  for x := 0 to 6 do
  begin
    StringGridCal.Cells[x, 1] := '';
  end;
  for x := 0 to 6 do
  begin
    StringGridCal.Cells[x, 5] := '';
  end;//////////////
  days[0] := 'Domingo';
  days[1] := 'Segunda';
  days[2] := 'Terça';
  days[3] := 'Quarta';
  days[4] := 'Quinta';
  days[5] := 'Sexta';
  days[6] := 'Sábado';
  with StringGridCal do
  begin
    for i := 0 to 6 do
       Cells[i, 0] := days[i]
  end;// pegar o número de dias
  iNumDays := MonthDays[IsLeapYear(Ano), Mes]; //Modifiquei aqui >
  formatsettings.ShortDateFormat := 'dd/mm/yyyy';
  iDay := DayOfWeek(StrToDate('01/'+IntToStr(Mes)+'/'+IntToStr(Ano)));  //E aqui para receber as variaveis Ano e Mês

  data:=datetostr(date);


  Labelcomp.Caption:= MESEXT(STRZERO(Mes,2))+'/'+anos;
  iRowCtr := 1;
  iColCtr := iDay - 1;
  for i := 1 to iNumDays do
  begin

    mes2:=FormatDateTime('mm', StrToDate('01/'+IntToStr(Mes)+'/'+IntToStr(Ano)) );
    anos:=FormatDateTime('yyyy', StrToDate('01/'+IntToStr(Mes)+'/'+IntToStr(Ano)));
    xtexto1:=strzero(I,2);
    xtexto1:=xtexto1+'/'+mes2+'/'+anos;

    DM_Tabelas.qryagenda.Close;
    DM_Tabelas.qryAgenda.SQL.Clear;
    DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', StrToDate(xtexto1)))); //que é essa 27/11/2012 00:00:00
    DM_Tabelas.qryAgenda.Open;

    if DM_Tabelas.qryagenda.RecordCount>0 then
    begin
       StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' Evento Programado';
    end
    else
    begin
        StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i);
        // feriados moveis
        iano:=(CalcFeriadosMoveis(Ano) - 47);
        if xtexto1=datetostr(iano) then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Carnaval';
        iano:=CalcFeriadosMoveis(Ano) - 2;
        if xtexto1=datetostr(iano) then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Sexta-Feira Santa';
        iano:=CalcFeriadosMoveis(Ano);
        if xtexto1=datetostr(iano) then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' -         Páscoa';
        iano:=CalcFeriadosMoveis(Ano) + 60;
        if xtexto1=datetostr(iano) then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+ ' -  Corpus Christi';

        if xtexto1='01/01/' + Ano.ToString then
          StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' -  Confraternização Universal';

        // feriados fixos
        if xtexto1='21/04/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Tiradentes';

        if xtexto1='01/05/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' -  Dia do Trabalho';
        if xtexto1='07/09/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Independência do Brasil';
        if xtexto1='12/10/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Nossa Senhora Aparecida';
        if xtexto1='02/11/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' -  Finados';
        if xtexto1='15/11/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' - Proclamação da República';
        if xtexto1='25/12/' + Ano.ToString then
           StringGridCal.Cells[iColCtr, iRowCtr] := IntToStr(i)+' -  Natal';

    end;
    Inc(iColCtr);
    if iColCtr > 6 then
    begin
      iColCtr := 0;
      Inc(iRowCtr);
    end;
  end;
end;
procedure TFrmAgenda.ds_qryagendaStateChange(Sender: TObject);
begin
  if FrmAgenda_Mov<>nil then
  begin
  //   FrmAgenda_Mov.btnIncluir.Enabled:=DtSrc.State in [dsBrowse,DsInactive];
     FrmAgenda_Mov.Btnconfirmar.Enabled:=FrmAgenda.ds_qryagenda.State in[dsInsert,Dsedit];
     FrmAgenda_Mov.BtnCancelar.Enabled:=FrmAgenda.ds_qryagenda.State in [dsinsert,dsEdit];
     FrmAgenda_Mov.BtnExcluir.Enabled:=FrmAgenda.ds_qryagenda.State in [dsBrowse];
  end;
end;

procedure TFrmAgenda.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DM_Tabelas.qryagenda.Close;
end;

procedure TFrmAgenda.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
  mes2:=FormatDateTime('mm', Date );
  anos:=FormatDateTime('yyyy', Date );
  Dias(mes2.ToInteger,anos.ToInteger); //Ano e mês desejado para preencher o grid assim que iniciar.
end;

procedure TFrmAgenda.FormShow(Sender: TObject);
var
vLinha,vColuna,I,y:integer;
xtexto1:String;

begin
   mes2:=FormatDateTime('mm', Date );
   anos:=FormatDateTime('yyyy', Date );
   Meses:=mes2.ToInteger;
   Dias(Meses,anos.ToInteger); //Agora ainda terá de decidir como será definido o ano.//Poderá utilizar uma varaivel no lugar onde está escrito 2015 para passar o ano para//o procedimento que escreverá no grid.
   Labelcomp.Caption:= MESEXT(STRZERO(Mes2.ToInteger,2))+'/'+anos;
   xtexto1:=copy(datetostr(date),1,2);
   xtexto1:=xtexto1+'/'+mes2+'/'+anos;
  //listar os horários daquela data
//   DM_Tabelas.qryAgenda.Close;
//   DM_Tabelas.qryAgenda.SQL.Clear;
//   DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', StrToDate(xtexto1)))); //que é essa 27/11/2012 00:00:00
//   DM_Tabelas.qryAgenda.Open;
  //zerar o tempo do MonthCalendar
//  DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', DateOf(MonthCalendar1.Date))));

    DM_Tabelas.qryAgenda.Close;
    DM_Tabelas.qryAgenda.SQL.Clear;
    DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', date))); //que é essa 27/11/2012 00:00:00
    DM_Tabelas.qryAgenda.Open;
end;

procedure TFrmAgenda.StringGridCalDblClick(Sender: TObject);
begin
  if DM_Tabelas.qryAgenda.RecordCount>0 then
     DM_Tabelas.qryAgenda.Edit
  else
  begin
    DM_Tabelas.qryAgenda.Append;
    DM_Tabelas.qryagenda.FieldByName('data').AsDateTime:=StrToDate(xtexto);
    DM_Tabelas.qryagenda.FieldByName('hora').AsDateTime:=time;
  end;
  try
    if FrmAgenda_Mov = nil then
       FrmAgenda_Mov:=TFrmAgenda_Mov.create(self);
    AbrirModal(Self, FrmAgenda_Mov);
  finally
    FreeAndNil(FrmAgenda_Mov);
  end;

end;

procedure TFrmAgenda.StringGridCalDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
  var
  tm,d,hoje:string;
  ps,vLinha,vColuna,y,i:integer;
  sTexto:string;
begin


  if copy(data,1,1)='0' then
      hoje:=copy(data,2,1)
  else
      hoje:=copy(data,1,2);
      tm:=StringGridCal.Cells[ACol, ARow];
  if (length(tm)>0) and (length(tm)>=2) then
  begin
    if (hoje=copy((StringGridCal.Cells[ACol, ARow]),1,2)) then
    begin
      with StringGridCal.Canvas do
      begin
          Brush.color := $004080FF;
          Font.Color  := clWhite;
          Font.Style  := [fsbold];
      end;
    end
    else
    begin

      if (gdselected in state) then
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;//coloque a cor do colordialog
        StringGridCal.Canvas.Brush.Color:= clblue;
      end
      else
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;
        StringGridCal.Canvas.Brush.Color:=clWindowFrame;
      end;
      ps:=pos('Evento',tm);
      if ps>0 then
      begin
        with StringGridCal.Canvas do
        begin
            Brush.color := $00000077;
            Font.Color := clyellow;
            Font.Style := [fsbold];
        end;
      end;

      ps:=pos('-',tm);// feriados
      if ps>0 then
      begin
        with StringGridCal.Canvas do
        begin
            Brush.color := clRed;
            Font.Color := clyellow;
            Font.Style := [fsbold];
        end;
      end;


      if (ACol = 0) and (ARow >= 0) then
      begin

        with StringGridCal.Canvas do
        begin
            Brush.color := clRed;
            Font.Color := clyellow;
            Font.Style := [fsbold];
        end;
      end;

    end;

  end;
  if (length(tm)>0) and (length(tm)<2) then
  begin
    if (hoje=copy((StringGridCal.Cells[ACol, ARow]),1,1)) then
    begin
      with StringGridCal.Canvas do
      begin
          Brush.color := $004080FF;
          Font.Color  := clWhite;
          Font.Style  := [fsbold];
      end;
    end
    else
    begin

      if (gdselected in state) then
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;//coloque a cor do colordialog
        StringGridCal.Canvas.Brush.Color:= clblue;
      end
      else
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;
        StringGridCal.Canvas.Brush.Color:=clWindowFrame;
      end;
      if (ACol = 0) and (ARow >= 0) then
      begin

        with StringGridCal.Canvas do
        begin
            Brush.color := clRed;
            Font.Color := clyellow;
            Font.Style := [fsbold];
        end;
      end;

    end;

  end;

  if (length(tm)=0) and (length(tm)<2) then
  begin
    if (hoje=copy((StringGridCal.Cells[ACol, ARow]),1,2)) then
    begin
      with StringGridCal.Canvas do
      begin
          Brush.color := $004080FF;
          Font.Color  := clWhite;
          Font.Style  := [fsbold];
      end;
    end
    else
    begin

      if (gdselected in state) then
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;//coloque a cor do colordialog
        StringGridCal.Canvas.Brush.Color:= clblue;
      end
      else
      begin
        StringGridCal.Canvas.Font.Color:= clWhite;
        StringGridCal.Canvas.Brush.Color:=clWindowFrame;
      end;
      if (ACol = 0) and (ARow >= 0) then
      begin

        with StringGridCal.Canvas do
        begin
            Brush.color := clRed;
            Font.Color := clyellow;
            Font.Style := [fsbold];
        end;
      end;

    end;

  end;

  StringGridCal.Canvas.FillRect(Rect);
  sTexto := StringGridCal.Cells[ACol, ARow];
  StringGridCal.Canvas.TextRect(Rect, sTexto, [tfWordBreak,tfVerticalCenter,tfCenter]);
//  StringGridCal.Canvas.TextOut(Rect.Left+2,Rect.Top+2,StringGridCal.Cells[ACol,Arow]);
end;

procedure TFrmAgenda.StringGridCalSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
  var
//  xtexto:String;
  posi,x:integer;
begin
    xtexto:=StringGridCal.Cells[ACol, ARow];
    x:=length(xtexto);
    if (x>0) and (x>=2) then
    begin
        xtexto:=copy(xtexto,1,2)+'/'+mes2+'/'+anos;
        DM_Tabelas.qryAgenda.Close;
        DM_Tabelas.qryAgenda.SQL.Clear;
        DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', StrToDate(xtexto)))); //que é essa 27/11/2012 00:00:00
        DM_Tabelas.qryAgenda.Open;

    end
    else if (x>0) and (x<2) then
    begin
      xtexto:='0'+xtexto+'/'+mes2+'/'+anos;
      DM_Tabelas.qryAgenda.Close;
      DM_Tabelas.qryAgenda.SQL.Clear;
      DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', StrToDate(xtexto)))); //que é essa 27/11/2012 00:00:00
      DM_Tabelas.qryAgenda.Open;
    end
    else
    begin
      DM_Tabelas.qryAgenda.Close;
    end;

  //listar os horários daquela data
  //zerar o tempo do MonthCalendar
//  qryAgenda.SQL.Add('SELECT * FROM agendados WHERE dtagendamento = ' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', DateOf(MonthCalendar1.Date))));
end;


procedure TFrmAgenda.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
