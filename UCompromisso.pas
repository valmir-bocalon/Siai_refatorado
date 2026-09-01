unit UCompromisso;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DBCtrls, ExtCtrls, Grids, DBGrids,
  DB, DBClient, DBTables, CheckLst, RXCtrls, Mask, ToolEdit, RXDBCtrl, Spin, Calendar,
  RxRichEd,DBLocals, DBRichEd, rxDBRichEd, rxToolEdit, XBanner;

type
  TFrmCompromisso = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    StatusBar1: TStatusBar;
    ChkLBx: TCheckListBox;
    RcDbTx: TRxDBRichEdit;
    DtSrc: TDataSource;
    CdsCompromisso: TClientDataSet;
    CdsCompromissoREGISTRO: TIntegerField;
    CdsCompromissoDT_REGISTRO: TDateField;
    CdsCompromissoASSUNTO: TStringField;
    CdsCompromissoDSC_ASSUNTO: TMemoField;
    CdsCompromissoCONF_OK: TStringField;
    XBanner1: TXBanner;
    BitBtn5: TBitBtn;
    DBEdit3: TDBEdit;
    DBEdit2: TDBEdit;
    DBDateEdit1: TDBDateEdit;
    DBEdit1: TDBEdit;
    Calendario: TCalendar;
    BitBtn3: TBitBtn;
    BitBtn2: TBitBtn;
    CAno: TSpinEdit;
    CMes: TSpinEdit;
    Cdia: TSpinEdit;
    procedure BitBtn5Click(Sender: TObject);
    procedure CdiaChange(Sender: TObject);
    procedure CMesChange(Sender: TObject);
    procedure CAnoChange(Sender: TObject);
    procedure CalendarioChange(Sender: TObject);
    procedure ChkLBxClick(Sender: TObject);
    procedure ChkLBxClickCheck(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
  Function ProximoDiaUtil (dData : TDate) : TDate;
    { Public declarations }
   
  end;

var
  FrmCompromisso: TFrmCompromisso;
  Cdata:string;
  reg:integer;

implementation

uses UDMAgenda, UAgenda, uRuntimeFields;

{$R *.dfm}

procedure TFrmCompromisso.BitBtn5Click(Sender: TObject);
VAR
I:Integer;
begin
   Cdata:=Cdia.Text+'/'+Cmes.Text+'/'+Cano.Text;
   cdata:=datetoStr(Proximodiautil(strtodate(Cdata)+1));
   IF ChkLBx.Items.Count > 0 then
   begin
      If Messagedlg('Existem compromissos n�o cumpridos,Deseja transferir para o pr�ximo dia util ?',
      mtConfirmation,[mbYes,mbNo],0)= mrYes then
      begin
         ChkLBx.ItemIndex:=0;
         for I:=1 to ChkLBx.Items.Count do
         begin
//            ChkLBx.ItemIndex:=0;
             Pesq_Comp.RecNo:=CHkLbx.ItemIndex+1;
             Cdscompromisso.Close;
             CdsCompromisso.Params[0].asinteger:=
             Pesq_CompREGISTRO.AsInteger;
             CdsCompromisso.Open;
            DtSrc.DataSet.edit;
            DbDateEdit1.Date:=strtoDate(cdata); // verificar a data
            DtSrc.DataSet.Post;
            (DtSrc.DataSet as TClientDataSet).ApplyUpdates(0);
            ChkLBx.ItemIndex:=ChkLBx.ItemIndex+1;
         end;
         close;
      end
      else
         close;
   end
   else
   close;
end;

procedure TFrmCompromisso.CdiaChange(Sender: TObject);
begin
   Calendario.Day:=cdia.Value;
end;

procedure TFrmCompromisso.CMesChange(Sender: TObject);
begin
   Calendario.Month:=Cmes.Value;
end;

procedure TFrmCompromisso.CAnoChange(Sender: TObject);
begin
   Calendario.Year:=Cano.Value;
end;

procedure TFrmCompromisso.CalendarioChange(Sender: TObject);
var
I:Integer;
begin
   Cdia.value:=Calendario.Day;
   Cmes.Value:=Calendario.Month;
   Cano.Value:=Calendario.Year;
   Cdata:=Cdia.Text+'/'+Cmes.Text+'/'+Cano.Text;
   ChkLBx.Clear;
   With DtSrcPesq.DataSet as TSQLClientDataSet do
   begin
      Close;
      Params[0].AsDate:=StrToDate(Cdata);
      open;
      if isEmpty then
        MessageDlg('Nenhum compromisso para esta data ',mtinformation,[mbOk],0)
      else
      for I:=1 to  Pesq_Comp.RecordCount do
      begin
         ChkLBx.Items.Add( Pesq_CompASSUNTO.AsString);
          Pesq_Comp.Next;
      end;
   end;
end;
procedure TFrmCompromisso.ChkLBxClick(Sender: TObject);
begin
    Pesq_Comp.RecNo:=chkLBx.ItemIndex+1;
   RcDbtx.DataField:='dsc_assunto';
   //reg:= Pesq_CompREGISTRO.AsInteger;
end;

procedure TFrmCompromisso.ChkLBxClickCheck(Sender: TObject);
begin
   If MessageDlg('Confirma o Compromisso cumprido ?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
   begin
       Pesq_Comp.RecNo:=CHkLbx.ItemIndex+1;
       begin
          Cdscompromisso.Close;
          CdsCompromisso.Params[0].asinteger:=
          Pesq_CompREGISTRO.AsInteger;
          CdsCompromisso.Open;
         DtSrc.DataSet.edit;
         Dbedit3.Text:='S';
         DtSrc.DataSet.Post;
         (DtSrc.DataSet as TClientDataSet).ApplyUpdates(0);
      end;
      ChkLBx.Items.Delete(ChkLBX.ItemIndex);
   end
   else
      ChkLbx.Checked[ChkLBx.ItemIndex]:=False;

end;

procedure TFrmCompromisso.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
   Cdia.Value:=Calendario.Day;
   Cmes.Value:=Calendario.Month;
   CAno.Value:=Calendario.Year;
end;

procedure TFrmCompromisso.BitBtn2Click(Sender: TObject);
var I:Integer;
begin
   Cdata:=Cdia.Text+'/'+Cmes.Text+'/'+Cano.Text;
   CdsCompromisso.Open;
   CdsCompromisso.Append;
   FrmAgenda:= TFrmAgenda.Create(self);
   try
      FrmAgenda.ShowModal;
   finally
      FrmAgenda.Free;
   end;
   ChkLBx.Clear;
   With DtSrcPesq.DataSet as TSQLClientDataSet do
   begin
      Close;
      Params[0].AsDate:=StrToDate(Cdata);
      open;
      if isEmpty then
        MessageDlg('Nenhum compromisso para esta data ',mtinformation,[mbOk],0)
      else
      for I:=1 to  Pesq_Comp.RecordCount do
      begin
         ChkLBx.Items.Add( Pesq_CompASSUNTO.AsString);
          Pesq_Comp.Next;
      end;
   end;
end;

procedure TFrmCompromisso.BitBtn3Click(Sender: TObject);
var I:integer;
begin
    CdsCompromisso.Close;
    CdsCompromisso.Params[0].Value:=
    Pesq_CompREGISTRO.Value;
    CdsCompromisso.Open;
   FrmAgenda:=TfrmAgenda.Create(self);
   try
      FrmAgenda.ShowModal;
   finally
      FrmAgenda.Free;
   end;
   ChkLBx.Clear;
   With DtSrcPesq.DataSet as TSQLClientDataSet do
   begin
      Close;
      Params[0].AsDate:=StrToDate(Cdata);
      open;
      if isEmpty then
        MessageDlg('Nenhum compromisso para esta data ',mtinformation,[mbOk],0)
      else
      for I:=1 to  Pesq_Comp.RecordCount do
      begin
         ChkLBx.Items.Add( Pesq_CompASSUNTO.AsString);
          Pesq_Comp.Next;
      end;
   end;
end;

function TFrmCompromisso.ProximoDiaUtil(dData: TDate): TDate;
begin
if DayOfWeek(dData) = 7 then
dData := dData + 2
else
if DayOfWeek(dData) = 1 then
dData := dData + 1;
ProximoDiaUtil := dData;
end;


procedure TFrmCompromisso.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
