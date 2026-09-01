unit UAgenda_mov;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, RxRichEd, DBCtrls, Mask,
  RXDBCtrl,DBClient, DB, ComCtrls, ExtCtrls, rxDBRichEd, rxToolEdit,
  XBanner, ACBrBase, ACBrEnterTab;

type
  TFrmAgenda_Mov = class(TForm)
    StatusBar1: TStatusBar;
    btnConfirmar: TBitBtn;
    BtnCancelar: TBitBtn;
    btnSair: TBitBtn;
    Panel1: TPanel;
    XBanner1: TXBanner;
    Label4: TLabel;
    Label2: TLabel;
    DBDateEdit1: TDBDateEdit;
    DBEdit1: TDBEdit;
    Label3: TLabel;
    Label1: TLabel;
    DBText1: TDBText;
    XBanner2: TXBanner;
    Label5: TLabel;
    DBEdit2: TDBEdit;
    Panel2: TPanel;
    BtnExcluir: TBitBtn;
    ACBrEnterTab1: TACBrEnterTab;
    DBRichEdit1: TDBRichEdit;
    procedure btnSairClick(Sender: TObject);
    procedure BtnExcluirClick(Sender: TObject);
    procedure btnConfirmarClick(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);
    procedure DBRichEdit1Enter(Sender: TObject);
    procedure DBRichEdit1Exit(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmAgenda_Mov: TFrmAgenda_Mov;

implementation

uses tabelas,Uagenda,funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrmAgenda_Mov.btnSairClick(Sender: TObject);
begin
  DM_Tabelas.qryAgenda.cancel;
  FrmAgenda.Dias(Meses,anos.ToInteger);
//  FrmAgenda.StringGridCal.repaint;
//  FrmAgenda.StringGridCal.Refresh;
  Close;
end;

procedure TFrmAgenda_Mov.DBRichEdit1Enter(Sender: TObject);
begin
  FrmAgenda_Mov.KeyPreview:=false;
end;

procedure TFrmAgenda_Mov.DBRichEdit1Exit(Sender: TObject);
begin
  FrmAgenda_Mov.KeyPreview:=true;
end;

procedure TFrmAgenda_Mov.FormShow(Sender: TObject);
begin
  DBEdit2.SetFocus;
end;

procedure TFrmAgenda_Mov.BtnExcluirClick(Sender: TObject);
begin
    If simnao('Deseja realmente excluir este registro ?','SIM') then
   begin
      DM_Tabelas.qryAgenda.cancel;
      DM_Tabelas.qryAgenda.delete;
   end;
end;

procedure TFrmAgenda_Mov.btnConfirmarClick(Sender: TObject);
begin
   if FrmAgenda.ds_qryagenda.State in [dsInsert,Dsedit] then
      DM_Tabelas.qryAgenda.Post;
end;

procedure TFrmAgenda_Mov.BtnCancelarClick(Sender: TObject);
begin
   DM_Tabelas.qryAgenda.cancel;
end;


procedure TFrmAgenda_Mov.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

end.
