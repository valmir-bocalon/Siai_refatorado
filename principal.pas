unit principal;

interface

uses
//  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
//  Dialogs,   dxButton, XBanner, ComCtrls, TabNotBk, ExtCtrls, IniFiles,
//  StdCtrls, Grids, DBGrids, Mask, DBCtrls, wwdbdatetimepicker, DB, ComCtrls, ImgList;
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, DBGrids, DB, StdCtrls, DBCtrls, ExtCtrls, jpeg, XBanner, Mask,
  XEdit, ComCtrls, IniFiles, ImgList, dxButton, Shellapi, Registry,
  JvExControls, JvXPCore, JvXPBar, TFlatHintUnit,qrprntr, TFlatPanelUnit,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, midaslib, dxCore2,psapi,RLConsts,
  System.ImageList, dxCoreA, Vcl.Buttons,  DateUtils,
  Notification,
  System.JSON,IdSSLOpenSSL,
  IdHTTP, acPNG;
 //RLConsts por causa da versao o fortes report no final das linhas tem mais
type
  TFrm_principal = class(TForm)
    ImageList1: TImageList;
    StatusBar1: TStatusBar;
    Timer1: TTimer;


    Image1: TImage;
    FlatHint1: TFlatHint;
    Panel1: TPanel;
    MainMenu1: TMainMenu;
    Vendas1: TMenuItem;
    Imveis1: TMenuItem;
    Financeiro1: TMenuItem;
    Bancos1: TMenuItem;
    ContaBancria1: TMenuItem;
    MovimentoBancrio1: TMenuItem;
    Imveis2: TMenuItem;
    Corretor1: TMenuItem;
    Usurio1: TMenuItem;
    Diversos1: TMenuItem;
    Participantes1: TMenuItem;
    Empreendimentos1: TMenuItem;
    iposdeDocumentos1: TMenuItem;
    PlanosdeContas1: TMenuItem;
    Cidades1: TMenuItem;
    Empresas1: TMenuItem;
    Configuraes1: TMenuItem;
    Utilitrios1: TMenuItem;
    Estrutura1: TMenuItem;
    Backup1: TMenuItem;
    AcertaVendaemRecebimento1: TMenuItem;
    Fechar1: TMenuItem;
    InibirExibirBarradeBotes1: TMenuItem;
    ContasReceberPagar1: TMenuItem;
    ReajustedeParcelas1: TMenuItem;
    Aditamento1: TMenuItem;
    Quitao1: TMenuItem;
    CessoTransferncia1: TMenuItem;
    abelaPrice1: TMenuItem;
    AcertaCdigodeOrigememRecebimento1: TMenuItem;
    AcertaPrticipanteXRecebimento1: TMenuItem;
    abelaIGPM1: TMenuItem;
    AcertaDatadoPrximoReajuste1: TMenuItem;
    AcertaroValorFinaldaVenda1: TMenuItem;
    ExcluirLoteamento1: TMenuItem;
    ColocardataeHoradoAditamentonamesmaLinha1: TMenuItem;
    AlteraroCampoPrximoReajuste1: TMenuItem;
    AdicionarMsAnodeReajuste1: TMenuItem;
    Cadastro1: TMenuItem;
    Email1: TMenuItem;
    ZQErro_Baixa: TZQuery;
    DS_ZQErro_Baixa: TDataSource;
    ConsultaRpida1: TMenuItem;
    LimparTabelaRemessaRecebidaporPerodo1: TMenuItem;
    xusuario: TLabel;
    Button1: TButton;
    lblvnd: TLabel;
    Timer2: TTimer;

    JXBarraCad_Diversos: TJvXPBar;
    JBarraUtil: TJvXPBar;
    A1: TMenuItem;
    A2: TMenuItem;
    E1: TMenuItem;
    P1: TMenuItem;
    D1: TMenuItem;
    T1: TMenuItem;
    toolbar: TPanel;
    dxButton26: TSpeedButton;
    dxButton24: TSpeedButton;
    dxButton22: TSpeedButton;
    dxButton25: TSpeedButton;
    dxButton3: TSpeedButton;
    dxButton20: TSpeedButton;
    dxButton1: TSpeedButton;
    dxButton14: TSpeedButton;
    dxButton28: TSpeedButton;
    BtFechar: TSpeedButton;
    dxButton23: TSpeedButton;
    dxButton2: TSpeedButton;
    dxButton19: TSpeedButton;
    dxButton18: TSpeedButton;
    dxButton27: TSpeedButton;
    Timer3: TTimer;
    A3: TMenuItem;
    Panel6: TPanel;
    Shape3: TShape;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    xdolar: TLabel;
    xeuro: TLabel;
    xbtc: TLabel;

    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure dxButton14Click(Sender: TObject);
    procedure BtFecharClick(Sender: TObject);
    procedure dxButton23Click(Sender: TObject);
    procedure dxButton27Click(Sender: TObject);
    procedure dxButton28Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure dxButton25Click(Sender: TObject);
    procedure dxButton24Click(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure dxButton20Click(Sender: TObject);
    procedure JXBarraRelItems0Click(Sender: TObject);
    procedure JXBarraCad_DiversosExit(Sender: TObject);
    procedure JXBarraCad_DiversosItems2Click(Sender: TObject);
    procedure JvXPBar1Items0Click(Sender: TObject);
    procedure JBarraUtilExit(Sender: TObject);
    procedure JXBarraCad_DiversosItems1Click(Sender: TObject);
    procedure JXBarraCad_DiversosItems3Click(Sender: TObject);
    procedure dxButton26Click(Sender: TObject);
    procedure JXBarraCad_DiversosItems4Click(Sender: TObject);
    procedure dxButton22Click(Sender: TObject);
    procedure JBarraUtilItems2Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Vendas1Click(Sender: TObject);
    procedure Imveis2Click(Sender: TObject);
    procedure Corretor1Click(Sender: TObject);
    procedure Usurio1Click(Sender: TObject);
    procedure Empreendimentos1Click(Sender: TObject);
    procedure ContaBancria1Click(Sender: TObject);
    procedure PlanosdeContas1Click(Sender: TObject);
    procedure Cidades1Click(Sender: TObject);
    procedure Empresas1Click(Sender: TObject);
    procedure Configuraes1Click(Sender: TObject);
    procedure AcertaVendaemRecebimento1Click(Sender: TObject);
    procedure Fechar1Click(Sender: TObject);
    procedure InibirExibirBarradeBotes1Click(Sender: TObject);
    procedure ContasReceberPagar1Click(Sender: TObject);
    procedure ReajustedeParcelas1Click(Sender: TObject);
    procedure Aditamento1Click(Sender: TObject);
    procedure Quitao1Click(Sender: TObject);
    procedure CessoTransferncia1Click(Sender: TObject);
    procedure Financeiro1Click(Sender: TObject);
    procedure abelaPrice1Click(Sender: TObject);
    procedure dxButton2Click(Sender: TObject);
    procedure AcertaCdigodeOrigememRecebimento1Click(Sender: TObject);
    procedure AcertaPrticipanteXRecebimento1Click(Sender: TObject);
    procedure abelaIGPM1Click(Sender: TObject);
    procedure AcertaDatadoPrximoReajuste1Click(Sender: TObject);
    procedure iposdeDocumentos1Click(Sender: TObject);
    procedure AcertaroValorFinaldaVenda1Click(Sender: TObject);
    procedure ExcluirLoteamento1Click(Sender: TObject);
    procedure ColocardataeHoradoAditamentonamesmaLinha1Click(
      Sender: TObject);
    procedure AlteraroCampoPrximoReajuste1Click(Sender: TObject);
    procedure JBarraUtilItems3Click(Sender: TObject);
    procedure AdicionarMsAnodeReajuste1Click(Sender: TObject);
    procedure Email1Click(Sender: TObject);
    procedure Cadastro1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ConsultaRpida1Click(Sender: TObject);
    procedure LimparTabelaRemessaRecebidaporPerodo1Click(Sender: TObject);
    procedure dxButton3Click(Sender: TObject);
    procedure dxButton19Click(Sender: TObject);
    procedure JXBarraCad_DiversosItems0Click(Sender: TObject);
    procedure JBarraUtilItems0Click(Sender: TObject);
    procedure JBarraUtilItems1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure A1Click(Sender: TObject);
    procedure A2Click(Sender: TObject);
    procedure P1Click(Sender: TObject);
    procedure D1Click(Sender: TObject);
    procedure T1Click(Sender: TObject);
    procedure JBarraUtilItems4Click(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);
    procedure A3Click(Sender: TObject);
    procedure JXBarraCad_DiversosItems5Click(Sender: TObject);
    procedure MenuItemDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; Selected: Boolean);
    procedure MenuItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
      var Width, Height: Integer);
  private
    { Private declarations }
    procedure AplicarEstiloMenu(AItem: TMenuItem);
    procedure ConfigurarMenuModerno;

    procedure AfterConstruction; override;

  public
    { Public declarations }
    procedure Prev(Sender: TObject);
    procedure GetCurrenciesQuotation;
  end;

var
  Frm_principal: TFrm_principal;
  VarUsuario   : string;
  VarCodUsuario: integer;
  senhaestrut  : Boolean;

  OldWidth : Integer = 0;
  OldHeight : Integer;


implementation

uses
  tabelas, funcoes, Cidade, Usuario, Loteamento,
  Estrutura, Empresa, Participante, Imoveis, Corretor, Cad_TipoDoc,
  UPlanoDeContas, UFrmContasBancarias, Venda, Configuracoes,
  Cad_Recebimento, Transferencia, TrocaCPF_RG, resposavel_conjuge,
  acertavenda_em_recebimento, Apresentacao, Aditamento, UFrmVisualizar,
  ReajusteDeParcelas, quitacao, Cessao, Price, acertaorigem_em_recebimento,
  acerta_adversa_recebimento, igpm, acertaproximo_reajuste_em_recebimento,
  acertavalor_final_venda, Excluir_loteamento, acerta_adiamento,
  ProximoReajusteDeParcelas, Acerto_ano_reajuste, Participante2,
  acerta_remessa_recebida,alerta, Acerto_parcelas_reajuste,
  ProximoVectoParcelas, Ucobranca, UFrmPainel, UFrmDash, ipca,
  zerarNossoNumero_em_recebimento, UAgenda, uRuntimeFields;

{$R *.dfm}


//current memory size of the current process in bytes  colocar no uses psapi
function CurrentMemoryUsage: Cardinal;
var   pmc: TProcessMemoryCounters;
begin
  pmc.cb := SizeOf(pmc) ;
  if GetProcessMemoryInfo(GetCurrentProcess, @pmc, SizeOf(pmc)) then
     Result := pmc.WorkingSetSize
  else
     RaiseLastOSError;
end;

Function VersaoExe: String;
type
   PFFI = ^vs_FixedFileInfo;
var
   F       : PFFI;
   Handle  : Dword;
   Len     : Longint;
   Data    : Pchar;
   Buffer  : Pointer;
   Tamanho : Dword;
   Parquivo: Pchar;
   Arquivo : String;
begin
   Arquivo  := Application.ExeName;
   Parquivo := StrAlloc(Length(Arquivo) + 1);
   StrPcopy(Parquivo, Arquivo);
   Len := GetFileVersionInfoSize(Parquivo, Handle);
   Result := '';
   if Len > 0 then
   begin
      Data:=StrAlloc(Len+1);
      if GetFileVersionInfo(Parquivo,Handle,Len,Data) then
      begin
         VerQueryValue(Data, '',Buffer,Tamanho);
         F := PFFI(Buffer);
         Result := Format('%d.%d.%d.%d',
                          [HiWord(F^.dwFileVersionMs),
                           LoWord(F^.dwFileVersionMs),
                           HiWord(F^.dwFileVersionLs),
                           Loword(F^.dwFileVersionLs)]
                         );
      end;
      StrDispose(Data);
   end;
   StrDispose(Parquivo);
end;
//resolução de tela
procedure PTela( Sender: Tobject );
  procedure PAlterar( Width, Height : Integer);
  var sDispMode: DevMode;
  begin
    EnumDisplaySettings(Nil,0,sDispMode);

    sDispMode.dmPelsWidth := Width;
    sDispMode.dmPelsHeight := Height;
    sDispMode.dmFields := DM_PELSWIDTH+DM_PELSHEIGHT;
    try
      ChangeDisplaySettings(sDispMode,0);
    except
      ShowMessage('Não é possivel alterar configurações de vídeo.')
    end;

  end;

begin

  if ((Screen.Width <> OldWidth ) or (Screen.Height <> OldHeight )) and (Oldwidth <> 0) then
    Palterar(OldWidth, OldHeight);

  // Aqui testo se a resolução é 800 porque quero utilizar pro meu software resolução 1280 X 768.
  If ((Screen.Width <> 1280 ) or ( Screen.Height <> 768)) then
  begin
{    if (messageDlg('Para uma boa utilização recomenda-se resolção de video de 1280 X 768. ' +
                    #13#13 + ' Deseja Alterar suas confirações de vídeo agora ?',
       mtConfirmation, [mbYes, mbNo], 0)=mrYes) then}
    begin
      // Vairavel que pega o valor original de inicialização ex: 1280 X 720... etc...
      OldWidth := GetSystemMetrics(SM_CXSCREEN);
      OldHeight := GetSystemMetrics(SM_CYSCREEN);
      // Resolução para qual quero utilizar enquanto o programa estiver aberto…
      if DM_Tabelas<>nil then
      begin
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '5120x2160'  then
            PAlterar(5120, 2160); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '3840x2160'  then
           PAlterar(1366, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '2560x1440'  then
            PAlterar(2560, 1440); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '2040x1080'  then
            PAlterar(2040, 1080); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString=  '2040x858' then
            PAlterar(2040, 858); // 1366 X 768
        if  DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1998x1080' then
            PAlterar(1998, 1080); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1920x1080' then
            PAlterar(1920, 1080); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString= '1680x1050'  then
            PAlterar(1680, 1050); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1600x900' then
           PAlterar(1600, 900); // 1600 X 900
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1366x768' then
           PAlterar(1366, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1360x768' then
           PAlterar(1360, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x1024' then
           PAlterar(1280,1024); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x960' then
           PAlterar(1280, 960); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x800' then
           PAlterar(1280, 800); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x768' then
           PAlterar(1280, 768); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x720' then
           PAlterar(1280, 720); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1280x600' then
           PAlterar(1280, 600); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='1024x768' then
           PAlterar(1024, 600); // 1366 X 768
        if DM_Tabelas.ZQConfiguracoes.FieldByName('resolucao_tela').AsString='800x600' then
           PAlterar(800, 600); // 1366 X 768
      end;
    end;
  end;

end;


procedure TFrm_principal.Prev(Sender: TObject);
begin
  if FrmVisualizar=nil then
     FrmVisualizar:=TFrmVisualizar.Create(Application);
  FrmVisualizar.QRPreview1.QRPrinter := TQRPrinter(Sender);
 FrmVisualizar.Show;
end;

procedure TFrm_principal.AplicarEstiloMenu(AItem: TMenuItem);
begin
  if AItem = nil then
    Exit;

  AItem.OnDrawItem := MenuItemDrawItem;
  AItem.OnMeasureItem := MenuItemMeasureItem;
end;

procedure TFrm_principal.ConfigurarMenuModerno;
var
  I: Integer;
begin
  MainMenu1.OwnerDraw := True;
  for I := 0 to MainMenu1.Items.Count - 1 do
    AplicarEstiloMenu(MainMenu1.Items[I]);
  DrawMenuBar(Handle);
end;

procedure TFrm_principal.MenuItemMeasureItem(Sender: TObject;
  ACanvas: TCanvas; var Width, Height: Integer);
var
  MenuItem: TMenuItem;
  IsTopLevel: Boolean;
  ShortcutText: string;
  CaptionText: string;
  MeasureRect: TRect;
begin
  MenuItem := Sender as TMenuItem;
  IsTopLevel := (MenuItem.Parent <> nil) and (MenuItem.Parent = MainMenu1.Items);

  if MenuItem.Caption = '-' then
  begin
    Width := 24;
    Height := 10;
    Exit;
  end;

  ACanvas.Font.Name := 'Segoe UI';
  ACanvas.Font.Style := [];
  if IsTopLevel then
    ACanvas.Font.Style := [fsBold];

  ShortcutText := '';
  if MenuItem.ShortCut <> 0 then
    ShortcutText := ShortCutToText(MenuItem.ShortCut);
  CaptionText := StringReplace(MenuItem.Caption, '&', '', [rfReplaceAll]);
  if CaptionText = '' then
    CaptionText := MenuItem.Caption;
  MeasureRect := Rect(0, 0, 0, 0);
  DrawText(ACanvas.Handle, PChar(CaptionText), -1, MeasureRect,
    DT_SINGLELINE or DT_NOPREFIX or DT_CALCRECT);

  if IsTopLevel then
  begin
    Width := (MeasureRect.Right - MeasureRect.Left) + 38;
    Height := 34;
  end
  else
  begin
    Width := (MeasureRect.Right - MeasureRect.Left) + 60;
    if ShortcutText <> '' then
      Inc(Width, ACanvas.TextWidth(ShortcutText) + 28);
    if MenuItem.Count > 0 then
      Inc(Width, 18);
    Height := 30;
  end;
end;

procedure TFrm_principal.MenuItemDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
const
  CBarColor = $00F4F6F8;
  CPopupColor = $00FCFCFD;
  CSelectedColor = $00E2ECF7;
  CAccentColor = $004F86F7;
  CTextColor = $002F343B;
  CDisabledColor = $00939AA3;
  CGutterColor = $00F0F2F5;
var
  MenuItem: TMenuItem;
  IsTopLevel: Boolean;
  TextRect: TRect;
  ShortcutText: string;
  CaptionText: string;

begin
  MenuItem := Sender as TMenuItem;
  IsTopLevel := (MenuItem.Parent <> nil) and (MenuItem.Parent = MainMenu1.Items);
  ShortcutText := '';
  if MenuItem.ShortCut <> 0 then
    ShortcutText := ShortCutToText(MenuItem.ShortCut);
  CaptionText := StringReplace(MenuItem.Caption, '&', '', [rfReplaceAll]);
  if CaptionText = '' then
    CaptionText := MenuItem.Caption;

  ACanvas.Brush.Style := bsSolid;
  if IsTopLevel then
    ACanvas.Brush.Color := CBarColor
  else
    ACanvas.Brush.Color := CPopupColor;

  if Selected then
    ACanvas.Brush.Color := CSelectedColor;

  ACanvas.FillRect(ARect);
  SetBkMode(ACanvas.Handle, TRANSPARENT);

  if MenuItem.Caption = '-' then
  begin
    ACanvas.Pen.Color := $00D9DEE5;
    ACanvas.MoveTo(ARect.Left + 10, (ARect.Top + ARect.Bottom) div 2);
    ACanvas.LineTo(ARect.Right - 10, (ARect.Top + ARect.Bottom) div 2);
    Exit;
  end;

  ACanvas.Font.Name := 'Segoe UI';
  ACanvas.Font.Style := [];
  ACanvas.Font.Color := CTextColor;
  if IsTopLevel then
    ACanvas.Font.Style := [fsBold];
  if not MenuItem.Enabled then
    ACanvas.Font.Color := CDisabledColor;

  if not IsTopLevel then
  begin
    ACanvas.Brush.Color := CGutterColor;
    ACanvas.FillRect(Rect(ARect.Left, ARect.Top, ARect.Left + 34, ARect.Bottom));
    if Selected then
    begin
      ACanvas.Brush.Color := CSelectedColor;
      ACanvas.FillRect(Rect(ARect.Left, ARect.Top, ARect.Left + 34, ARect.Bottom));
    end;
  end;

  if Selected then
  begin
    ACanvas.Pen.Color := CAccentColor;
    if IsTopLevel then
    begin
      ACanvas.MoveTo(ARect.Left + 2, ARect.Bottom - 2);
      ACanvas.LineTo(ARect.Right - 2, ARect.Bottom - 2);
    end
    else
    begin
      ACanvas.MoveTo(ARect.Left + 1, ARect.Top + 3);
      ACanvas.LineTo(ARect.Left + 1, ARect.Bottom - 3);
    end;
  end;

  if IsTopLevel then
  begin
    TextRect := Rect(ARect.Left + 14, ARect.Top, ARect.Right - 14, ARect.Bottom);
  end
  else
  begin
    TextRect := Rect(ARect.Left + 40, ARect.Top, ARect.Right - 14, ARect.Bottom);
  end;

  DrawText(ACanvas.Handle, PChar(CaptionText), -1, TextRect,
    DT_SINGLELINE or DT_VCENTER or DT_LEFT or DT_NOPREFIX);

  if (not IsTopLevel) and (ShortcutText <> '') then
  begin
    ACanvas.Font.Color := $00737C86;
    DrawText(ACanvas.Handle, PChar(ShortcutText), -1, TextRect,
      DT_SINGLELINE or DT_VCENTER or DT_RIGHT);
    if not MenuItem.Enabled then
      ACanvas.Font.Color := CDisabledColor;
  end;

  if (not IsTopLevel) and (MenuItem.Count > 0) then
  begin
    ACanvas.Font.Color := $00737C86;
    DrawText(ACanvas.Handle, PChar('>'), -1, TextRect,
      DT_SINGLELINE or DT_VCENTER or DT_RIGHT);
  end;
end;

procedure TFrm_principal.FormShow(Sender: TObject);
var
  barra:TIniFile;
  VarPath,VarExibe:String;
  BMP : TBitMap;
  JPG : TJpegImage;
  dia,mes,ano:integer;
  xano,data:string;
  sAno, sMes, sDia: Word;

begin
   try
     // Chama a procedure que busca as cotações
     GetCurrenciesQuotation;
     Panel6.Visible:=true;
   except
     Panel6.Visible:=false;
   end;
  if FrmApresentacao=nil then
     FrmApresentacao:=TFrmApresentacao.Create(Application);
  FrmApresentacao.Showmodal;


  DM_Tabelas.ZQUsuario.Open;
  DM_Tabelas.ZQUsuTemp.Open;
  DM_Tabelas.ZQLoc_Senha.Open;
  DM_Tabelas.ZQPermissoes.Open;
  DM_Tabelas.ZQLogin.Open;
  DM_Tabelas.ZQConfiguracoes.Open;
  DM_Tabelas.ZQEmpresa.Open;

  ZQErro_Baixa.ExecSQL;

  if not Verif_senha(' Principal','Entrar no sistema','') Then Close;
  if not CSenha Then Begin
    Close;
    exit
  end;

  DecodeDate( Date, sAno, sMes, sDia );
  xAno := IntToStr(SAno);
                                                      //alt+184   alt+0174
  StatusBar1.Panels.Items[3].Text := 'V.'+VersaoExe+'-'+xAno+' © Master Software ®  -  Usuário:'+xusuario.Caption;//+'  | Serial da Placa Mãe:'+GetMotherboardSerialNumber; // MemoNovidade.Lines.Strings[0];
  DM_Tabelas.ZQEmpresa.Open;
  if DM_TAbelas.ZQEmpresa.FieldByName('cad_empresa').AsDateTime<GetFileDate('SIAI.exe') then Begin
    DM_tabelas.ZQEmpresa.Edit;
    DM_TAbelas.ZQEmpresa.FieldByName('cad_empresa').AsDateTime:=GetFileDate('SIAI.exe');
    DM_TAbelas.ZQEmpresa.Post;
    showmessage('Registrada nova versão do sistema...');
  End
  else if DM_TAbelas.ZQEmpresa.FieldByName('cad_empresa').AsDateTime>GetFileDate('SIAI.exe') then Begin
    Showmessage('A data do executavel usado neste terminal é de '+datetostr(GetFileDate('SIAI.exe'))+chr(13)+chr(13)
    +'A data do executavel mais atual usado por esta empresa é de '+datetostr(DM_TAbelas.ZQEmpresa.FieldByName('cad_empresa').AsDateTime)
    +chr(13)+chr(13)+'Providencie uma atualização URGENTE, para evitar danos ao banco de dados!!')
  end;


                          DM_Tabelas.ZQCorretor.Open;
                          DM_Tabelas.ZQVendedor.Open;
                          DM_Tabelas.ZQQuadras.Open;
                          DM_Tabelas.ZQVenda.Open;
                          DM_Tabelas.ZQContaBancaria.Open;
                          DM_Tabelas.ZQEspecial.Open;
                          DM_Tabelas.ZQTipodoc.Open;
                          DM_Tabelas.ZQLoteamento.Open;
                          DM_Tabelas.ZQRemes_Receb.Open;




  //chama resolução de tela
   PTela(Sender);



  Panel1.Visible:=true;
  Application.ProcessMessages;

  Panel1.Visible:=false;


  data:=datetostr(date);
  dia:=strtoint(copy(data,1,2));
  mes:=strtoint(copy(data,4,2));
  ano:=strtoint(copy(data,7,4));
  StatusBar1.Panels.Items[0].Text := EstacaoDoAno(date)+'-'+CalendarioLunar(dia,mes,ano);
  timer3.Enabled:=true;

 // TThread.CreateAnonymousThread(procedure
        //begin
            try
              DM_Tabelas.qryAgenda.Close;
              DM_Tabelas.qryAgenda.SQL.Clear;
              DM_Tabelas.qryAgenda.SQL.Add('SELECT * FROM agenda WHERE data = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', date))); //que é essa 27/11/2012 00:00:00
              DM_Tabelas.qryAgenda.Open;
              if DM_Tabelas.qryAgenda.RecordCount>0 then
              begin
                if FrmAgenda = nil then
                   FrmAgenda:=TFrmAgenda.create(self);
                FrmAgenda.showmodal;
              end;
            finally
               DM_Tabelas.qryAgenda.Close;
               FreeAndNil(FrmAgenda);
            end;
//          end).start();
end;

procedure TFrm_principal.GetCurrenciesQuotation;
var
  HTTPClient: TIdHTTP;
  JSONResponse: string;
  JSONObj: TJSONObject;
  USD_Rate, EUR_Rate, BTC_Rate: string;
  SSLHandler: TIdSSLIOHandlerSocketOpenSSL;
begin
  HTTPClient := TIdHTTP.Create(nil);
  SSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  try
    SSLHandler.SSLOptions.Method     := sslvTLSv1_2;  // Força TLS 1.2
    SSLHandler.SSLOptions.Mode       := sslmClient;
    HTTPClient.IOHandler             := SSLHandler;
    HTTPClient.HandleRedirects       := True;  // Permite redirecionamentos
    HTTPClient.Request.UserAgent     := 'Mozilla/5.0';  // Alguns servidores rejeitam UserAgent vazio
    // Faz a requisição GET
    JSONResponse := HTTPClient.Get('https://economia.awesomeapi.com.br/json/last/USD-BRL,EUR-BRL,BTC-BRL');

    // Converte o response para JSON
    JSONObj := TJSONObject.ParseJSONValue(JSONResponse) as TJSONObject;
    try
      // Extrai os valores das cotações
      USD_Rate := JSONObj.GetValue<string>('USDBRL.bid');
      EUR_Rate := JSONObj.GetValue<string>('EURBRL.bid');
      BTC_Rate := JSONObj.GetValue<string>('BTCBRL.bid');

      // Exibe as cotações
      xdolar.Caption := USD_Rate;
      xeuro.Caption  := EUR_Rate;
      xbtc.Caption   := BTC_Rate;
    finally
      JSONObj.Free;
    end;
  finally
    HTTPClient.Free;
    SSLHandler.Free;
  end;
end;

procedure TFrm_principal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//  if not senhaestrut Then Begin
{    if DM_Tabelas.ZQLoc_Senha.Active=false then
       DM_Tabelas.ZQLoc_Senha.open;
    if DM_Tabelas.ZQLogin.Active=false then
       DM_Tabelas.ZQLogin.open;
    if DM_Tabelas.ZQTipodoc.Active=false then
       DM_Tabelas.ZQTipodoc.Open;
    if not Verif_senha(' Principal','Sair no sistema','') Then Exit;}
    DM_Tabelas.ZQAgrupaGrupo.Close;
    DM_Tabelas.ZQUsuario.Close;
    DM_Tabelas.ZQCidade.Close;
    DM_Tabelas.ZQLogin.Close;
    DM_Tabelas.ZQUsuTemp.Close;
    DM_Tabelas.ZQLoc_Senha.Close;
    DM_Tabelas.ZQPermissoes.Close;
    DM_Tabelas.ZQLoteamento.Close;
    DM_TAbelas.ZQEmpresa.Close;
    DM_Tabelas.ZqParticipante.Close;
    DM_Tabelas.ZQincorp_loteame.Close;
    DM_Tabelas.ZQContaBancaria.close;
    DM_Tabelas.ZQMens_Retorno.close;
    DM_Tabelas.ZQComprador.close;
    DM_Tabelas.ZQCorretor.close;
    DM_Tabelas.ZQVendedor.close;
    DM_Tabelas.ZQQuadras.close;
    DM_Tabelas.ZQVenda.close;
    DM_Tabelas.ZQEspecial.close;
    DM_Tabelas.ZQTipodoc.close;
    DM_Tabelas.ZQresponsavel.close;
    DM_Tabelas.ZQImovel.close;
    DM_Tabelas.ZQNumOrdem.close;
    DM_Tabelas.ZQRecebBxTemp.close;
    DM_Tabelas.ZQReBxHi.close;
    DM_Tabelas.ZQReceb_Baixa.close;
    DM_Tabelas.ZQRecebimento.Close;
    DM_Tabelas.ZQCobaRe.close;
    DM_Tabelas.ZQCheque.close;
    DM_Tabelas.ZQConfiguracoes.close;
    DM_Tabelas.ZQCaixa.close;
    DM_Tabelas.ZQMovBancaria.Close;
    DM_Tabelas.ZQPlanoDeContas.Close;
    DM_Tabelas.ZQBol_men.Close;
    DM_Tabelas.ZQRemes_Receb.Close;
    DM_Tabelas.ZQBancRemes.Close;
    DM_Tabelas.ZQProfissao.Close;
    DM_Tabelas.ZQConjuge.close;
    DM_Tabelas.ZQCompr_conjuge.Close;
    DM_Tabelas.ZQCompr_Dados.Close;
    DM_Tabelas.ZQCompr_Resp_Dados.Close;

//    DM_Tabelas.ZQlog_logradouro.close;
//    DM_Tabelas.zconeccao.Disconnect;
  //end;
end;

procedure TFrm_principal.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_principal.dxButton14Click(Sender: TObject);
begin
  if not Verif_senha('Usuário','Tela de Usuário','') Then Exit;
  try
    if Frm_Usuario=nil then
       Frm_Usuario := TFrm_Usuario.Create(self);
    Frm_Usuario.Left := Frm_principal.Left+7;
    Frm_Usuario.Top := Frm_principal.Top+135;
    Frm_Usuario.showModal;
  finally
    Frm_Usuario:=nil;
    Frm_Usuario.Free;
  end;
end;

procedure TFrm_principal.BtFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_principal.dxButton23Click(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Tela de Loteamento','') Then Exit;
//  Frm_Loteamento.Left := Frm_principal.Left+7;
//  Frm_Loteamento.Top := Frm_principal.Top+135;
//  Frm_Loteamento.show;
  try
    if Frm_Loteamento=nil then
       Frm_Loteamento := TFrm_Loteamento.Create(self);
    Frm_Loteamento.Left := Frm_principal.Left+7;
    Frm_Loteamento.Top := Frm_principal.Top+135;
    Frm_Loteamento.showModal;
  finally
    Frm_Loteamento:=nil;
    Frm_Loteamento.Free;
  end;

end;

procedure TFrm_principal.dxButton27Click(Sender: TObject);
begin
  JBarraUtil.Left := 756;
  JBarraUtil.Visible := True;
  JBarraUtil.SetFocus;
end;

procedure TFrm_principal.dxButton28Click(Sender: TObject);
begin
  JXBarraCad_Diversos.Visible := True;
  JXBarraCad_Diversos.SetFocus;
end;

procedure TFrm_principal.T1Click(Sender: TObject);
begin
  try
    if Frm_ipca=nil then
       Frm_ipca := TFrm_ipca.Create(self);
    Frm_ipca.Left := Frm_principal.Left+7;
    Frm_ipca.Top := Frm_principal.Top+135;
    Frm_ipca.showModal;
  finally
    Frm_ipca:=nil;
    Frm_ipca.Free;
  end;
end;

procedure TFrm_principal.Timer1Timer(Sender: TObject);
begin
  StatusBar1.Panels[2].Text := ' '+DateToStr(date)+' - '+diasemana(date)+' - '+TimeToStr(Time);
end;

procedure TFrm_principal.dxButton25Click(Sender: TObject);
begin
  if not Verif_senha('Participante','Tela Principal','') Then Exit;
//  Frm_participante.Left := Frm_principal.Left+7;
//  Frm_participante.Top := Frm_principal.Top+135;
//  Frm_Participante.show;

  try
    if Frm_Participante=nil then
       Frm_Participante := TFrm_Participante.Create(self);
//    Frm_Participante.Left := Frm_principal.Left+7;
//    Frm_Participante.Top := Frm_principal.Top+135;
   // Frm_Participante.showModal;
   Frm_Participante.show;
  finally
//    Frm_Participante:=nil;
//    Frm_Participante.Free;
  end;


end;

procedure TFrm_principal.dxButton24Click(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Tela Principal','') Then Exit;
//  Frm_Imoveis.Left := Frm_principal.Left+7;
//  Frm_Imoveis.Top := Frm_principal.Top+135;
//  Frm_Imoveis.show;
  try
    if Frm_Imoveis=nil then
       Frm_Imoveis := TFrm_Imoveis.Create(self);
    Frm_Imoveis.Left := Frm_principal.Left+7;
    Frm_Imoveis.Top := Frm_principal.Top+135;
    Frm_Imoveis.showModal;
  finally
    Frm_Imoveis:=nil;
    Frm_Imoveis.Free;
  end;

end;

procedure TFrm_principal.dxButton1Click(Sender: TObject);
begin
  if not Verif_senha('Corretor','Tela Principal','') Then Exit;
//  Frm_Corretor.Left := Frm_principal.Left+7;
//  Frm_Corretor.Top := Frm_principal.Top+135;
//  Frm_Corretor.show;

  try
    if Frm_Corretor=nil then
       Frm_Corretor := TFrm_Corretor.Create(self);
    Frm_Corretor.Left := Frm_principal.Left+7;
    Frm_Corretor.Top := Frm_principal.Top+135;
    Frm_Corretor.showModal;
  finally
    Frm_Corretor:=nil;
    Frm_Corretor.Free;
  end;

end;

procedure TFrm_principal.dxButton20Click(Sender: TObject);
begin
  try
    if FrmCad_ContasBancarias=nil then
       FrmCad_ContasBancarias := TFrmCad_ContasBancarias.Create(self);
    FrmCad_ContasBancarias.Left := Frm_principal.Left+7;
    FrmCad_ContasBancarias.Top := Frm_principal.Top+135;
    FrmCad_ContasBancarias.showModal;
  finally
    FrmCad_ContasBancarias:=nil;
    FrmCad_ContasBancarias.Free;
  end;
//
//  FrmCad_ContasBancarias.Left := Frm_principal.Left+7;
//  FrmCad_ContasBancarias.Top := Frm_principal.Top+135;
//  FrmCad_ContasBancarias.show;
end;

procedure TFrm_principal.JXBarraRelItems0Click(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documento','Tela Principal','') then exit;
  if FrmCad_TipoDoc=nil then
     FrmCad_TipoDoc:=TFrmCad_TipoDoc.Create(Self);
  FrmCad_TipoDoc.Left := Frm_principal.Left+7;
  FrmCad_TipoDoc.Top := Frm_principal.Top+135;
  FrmCad_TipoDoc.showmodal;
  FrmCad_TipoDoc:=nil;
  FrmCad_TipoDoc.Free;
end;

procedure TFrm_principal.JXBarraCad_DiversosExit(Sender: TObject);
begin
   JXBarraCad_Diversos.Visible := False;
end;

procedure TFrm_principal.JXBarraCad_DiversosItems2Click(Sender: TObject);
begin
  if not Verif_senha('Cidade','Tela de Cidade','') Then Exit;
//  Frm_Cidade.Left := Frm_principal.Left+7;
//  Frm_Cidade.Top := Frm_principal.Top+135;
//  Frm_Cidade.show;
  try
    if Frm_Cidade=nil then
       Frm_Cidade := TFrm_Cidade.Create(self);
    Frm_Cidade.Left := Frm_principal.Left+7;
    Frm_Cidade.Top := Frm_principal.Top+135;
    Frm_Cidade.showModal;
  finally
    Frm_Cidade:=nil;
    Frm_Cidade.Free;
    JXBarraCad_Diversos.Visible := False;
  end;



end;

procedure TFrm_principal.JvXPBar1Items0Click(Sender: TObject);
begin
//   FrmEstrutura.showmodal;
end;

procedure TFrm_principal.JBarraUtilExit(Sender: TObject);
begin
  JBarraUtil.Visible := False;
end;

procedure TFrm_principal.JXBarraCad_DiversosItems1Click(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Tela Principal','') then exit;
//  FrmCad_PlanodeContas.show;
  try
    if FrmCad_PlanodeContas=nil then
       FrmCad_PlanodeContas := TFrmCad_PlanodeContas.Create(self);
    FrmCad_PlanodeContas.Left := Frm_principal.Left+7;
    FrmCad_PlanodeContas.Top := Frm_principal.Top+135;
    FrmCad_PlanodeContas.showModal;
  finally
    FrmCad_PlanodeContas:=nil;
    FrmCad_PlanodeContas.Free;
  end;

  JXBarraCad_Diversos.Visible := False;
end;

procedure TFrm_principal.JXBarraCad_DiversosItems3Click(Sender: TObject);
begin
  if not Verif_senha('Empresa','Tela Principal','') then exit;
//  FrmCad_Empresa.Left := Frm_principal.Left+7;
//  FrmCad_Empresa.Top := Frm_principal.Top+135;
//  FrmCad_Empresa.show;

  try
    if FrmCad_Empresa=nil then
       FrmCad_Empresa := TFrmCad_Empresa.Create(self);
    FrmCad_Empresa.Left := Frm_principal.Left+7;
    FrmCad_Empresa.Top := Frm_principal.Top+135;
    FrmCad_Empresa.showModal;
  finally
    FrmCad_Empresa:=nil;
    FrmCad_Empresa.Free;
  end;


  JXBarraCad_Diversos.Visible := False;

end;

procedure TFrm_principal.dxButton26Click(Sender: TObject);
begin
  Panel1.Visible:=true;
  Application.ProcessMessages;
  //Frm_Venda.Left := Frm_principal.Left+7;
//  Frm_Venda.Top := Frm_principal.Top+135;

  if not Verif_senha('Venda','Tela Principal','') then exit;

{  if DM_tabelas.ZQContaBancaria.Active=false then
     DM_tabelas.ZQContaBancaria.open;
  if DM_tabelas.ZQQuadras.Active=false then
     DM_tabelas.ZQQuadras.open;
  if DM_Tabelas.ZQReBxHi.Active=false then
     DM_Tabelas.ZQReBxHi.open;
  if DM_Tabelas.ZQincorp_loteame.Active=false then
     DM_Tabelas.ZQincorp_loteame.Open;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  if DM_Tabelas.zqprocuradores.Active=false then
     DM_Tabelas.zqprocuradores.open;
  if DM_tabelas.ZQNumOrdem.Active = false then
     DM_tabelas.ZQNumOrdem.Open;
  if DM_Tabelas.ZQImovel.Active = false then
     DM_Tabelas.ZQImovel.open;
  if DM_Tabelas.ZQVendedor.Active=false then
     DM_Tabelas.ZQVendedor.open;
  if DM_Tabelas.ZQRecebimento.Active=false then
     DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZQVenda.Active=false then
     DM_Tabelas.ZQVenda.open;
  if DM_Tabelas.ZQComprador.Active=false then
     DM_Tabelas.ZQComprador.open;
  if DM_tabelas.ZQCheque.Active=false then
     DM_tabelas.ZQCheque.open;
  if DM_Tabelas.ZQTipoDoc.Active=false then
     DM_Tabelas.ZQTipoDoc.open;
  DM_Tabelas.ZQTipoDoc.First;   }
  Panel1.Visible:=false;
  Application.ProcessMessages;
  if Frm_Venda=nil then
     Frm_Venda:=TFrm_Venda.Create(Application);
  Frm_Venda.show;

end;

procedure TFrm_principal.JXBarraCad_DiversosItems4Click(Sender: TObject);
begin
  if not Verif_senha('Configurações','Tela Principal','') then exit;
  Frm_Configuracoes.Left := Frm_principal.Left+7;
  Frm_Configuracoes.Top := Frm_principal.Top+135;
  Frm_Configuracoes.show;
  JXBarraCad_Diversos.Visible := False;
end;

procedure TFrm_principal.JXBarraCad_DiversosItems5Click(Sender: TObject);
begin
  try
     JXBarraCad_Diversos.Visible := False;
    if FrmAgenda = nil then
       FrmAgenda:=TFrmAgenda.create(self);
    FrmAgenda.showmodal;
  finally
    FreeAndNil(FrmAgenda);
  end;
end;

procedure TFrm_principal.dxButton22Click(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Tela Principal','') then exit;
//  FrmCad_Recebimento.Left := Frm_principal.Left+7;
//  FrmCad_Recebimento.Top := Frm_principal.Top+135;
//  FrmCad_Recebimento.show;
  Panel1.Visible:=true;
  Application.ProcessMessages;


{  DM_tabelas.ZQTipodoc.close;
  DM_tabelas.ZQTipodoc.SQL.Clear;
  DM_tabelas.ZQTipodoc.sql.Add('select * from tipodocumento');
  DM_tabelas.ZQTipodoc.open;
  DM_tabelas.ZQTipodoc.First;

  DM_tabelas.ZQPlanoDeContas.close;
  DM_tabelas.ZQPlanoDeContas.SQL.clear;
  DM_tabelas.ZQPlanoDeContas.SQL.Add('Select * from Plano_contas');
  DM_tabelas.ZQPlanoDeContas.open;

  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist limit 0,1');
  DM_tabelas.ZQReBxHi.open;


  DM_TAbelas.ZQRecebBxTemp.Close;
  DM_TAbelas.ZQRecebBxTemp.SQL.Clear;
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('SELECT * FROM (recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib)join receb_baixa as B ON B.refbaixa=H.refer limit 0,1');
  DM_TAbelas.ZQRecebBxTemp.open;

  DM_Tabelas.ZQCheque.Close;
  DM_Tabelas.ZQCheque.SQL.Clear;
  DM_Tabelas.ZQCheque.SQL.Add('select * from cheque_rec');
  DM_Tabelas.ZQCheque.open;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento where quadralote is not null or quadralote !='''' order by DT_Vencimento limit 0,20');
  DM_tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQRecebimento.Last;

  DM_Tabelas.ZQContaBancaria.SQL.Clear;
  DM_Tabelas.ZQContaBancaria.SQL.Add('Select * from conta_bancaria');
  DM_Tabelas.ZQContaBancaria.Open;   }



  if FrmCad_Recebimento=nil then
     //FrmCad_Recebimento:=TFrmCad_Recebimento.Create(Application)
     FrmCad_Recebimento := TFrmCad_Recebimento.Create(Self)
  else
     FrmCad_Recebimento.bringToFront;
  FrmCad_Recebimento.ZQGerou.Open;
  FrmCad_Recebimento.Show;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_principal.JBarraUtilItems2Click(Sender: TObject);
begin
  JBarraUtil.Visible := False;
//  Frm_acertavenda_em_recebimento.showmodal;
 try
    if Frm_acertavenda_em_recebimento=nil then
       Frm_acertavenda_em_recebimento   := TFrm_acertavenda_em_recebimento.Create(self);
    Frm_acertavenda_em_recebimento.Left := Frm_principal.Left+7;
    Frm_acertavenda_em_recebimento.Top  := Frm_principal.Top+135;
    Frm_acertavenda_em_recebimento.showModal;
  finally
    Frm_acertavenda_em_recebimento:=nil;
    Frm_acertavenda_em_recebimento.Free;
  end;
end;

procedure TFrm_principal.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//if (ssCtrl in Shift) and (chr(Key) in ['A', 'a']) then
//  SHOWMESSAGE('');
end;

procedure TFrm_principal.Vendas1Click(Sender: TObject);
begin
  Panel1.Visible:=false;
  Application.ProcessMessages;

//  Frm_Venda.Left := Frm_principal.Left+7;
  //Frm_Venda.Top := Frm_principal.Top+135;
  if not Verif_senha('Venda','Tela Principal','') then exit;
{  if DM_tabelas.ZQContaBancaria.Active=false then
     DM_tabelas.ZQContaBancaria.open;
  if DM_tabelas.ZQQuadras.Active=false then
     DM_tabelas.ZQQuadras.open;
  if DM_Tabelas.ZQReBxHi.Active=false then
     DM_Tabelas.ZQReBxHi.open;
  if DM_Tabelas.ZQincorp_loteame.Active=false then
     DM_Tabelas.ZQincorp_loteame.Open;
  if DM_Tabelas.ZQLoteamento.Active=false then
     DM_Tabelas.ZQLoteamento.Open;
  if DM_Tabelas.zqprocuradores.Active=false then
     DM_Tabelas.zqprocuradores.open;
  if DM_tabelas.ZQNumOrdem.Active = false then
     DM_tabelas.ZQNumOrdem.Open;
  if DM_Tabelas.ZQImovel.Active = false then
     DM_Tabelas.ZQImovel.open;
  if DM_Tabelas.ZQVendedor.Active=false then
     DM_Tabelas.ZQVendedor.open;
  if DM_Tabelas.ZQRecebimento.active=false then
     DM_Tabelas.ZQRecebimento.open;
  if DM_Tabelas.ZQVenda.Active=false then
     DM_Tabelas.ZQVenda.open;
  if DM_Tabelas.ZQComprador.Active=false then
     DM_Tabelas.ZQComprador.open;
  if DM_tabelas.ZQCheque.Active=false then
     DM_tabelas.ZQCheque.open;
  if DM_Tabelas.ZQTipoDoc.Active=false then
     DM_Tabelas.ZQTipoDoc.open;
  DM_Tabelas.ZQTipoDoc.First;}
  if Frm_Venda=nil then
     Frm_Venda:=TFrm_Venda.Create(Application);
  Frm_Venda.show;

end;

procedure TFrm_principal.Imveis2Click(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Tela Principal','') Then Exit;
//  Frm_Imoveis.Left := Frm_principal.Left+7;
//  Frm_Imoveis.Top := Frm_principal.Top+135;
//  Frm_Imoveis.show;
  try
    if Frm_Imoveis=nil then
       Frm_Imoveis := TFrm_Imoveis.Create(self);
    Frm_Imoveis.Left := Frm_principal.Left+7;
    Frm_Imoveis.Top := Frm_principal.Top+135;
    Frm_Imoveis.showModal;
  finally
    Frm_Imoveis:=nil;
    Frm_Imoveis.Free;
  end;

end;

procedure TFrm_principal.Corretor1Click(Sender: TObject);
begin
  if not Verif_senha('Corretor','Tela Principal','') Then Exit;
//  Frm_Corretor.Left := Frm_principal.Left+7;
//  Frm_Corretor.Top := Frm_principal.Top+135;
//  Frm_Corretor.show;
  try
    if Frm_Corretor=nil then
       Frm_Corretor := TFrm_Corretor.Create(self);
    Frm_Corretor.Left := Frm_principal.Left+7;
    Frm_Corretor.Top := Frm_principal.Top+135;
    Frm_Corretor.showModal;
  finally
    Frm_Corretor:=nil;
    Frm_Corretor.Free;
  end;

end;

procedure TFrm_principal.D1Click(Sender: TObject);
begin
   if not Verif_senha('Estatíticas','DashBoard','') then exit;
  if FrmDash=nil then
     FrmDash:=TFrmDash.create(self);
  FrmDash.showmodal;
  FreeAndNil(FrmDash);
end;

procedure TFrm_principal.Usurio1Click(Sender: TObject);
begin
  if not Verif_senha('Usuário','Tela de Usuário','') Then Exit;
//  Frm_Usuario.Left := Frm_principal.Left+7;
//  Frm_Usuario.Top := Frm_principal.Top+135;
//  Frm_Usuario.show;
  try
    if Frm_Usuario=nil then
       Frm_Usuario := TFrm_Usuario.Create(self);
    Frm_Usuario.Left := Frm_principal.Left+7;
    Frm_Usuario.Top := Frm_principal.Top+135;
    Frm_Usuario.showModal;
  finally
    Frm_Usuario:=nil;
    Frm_Usuario.Free;
  end;

end;

procedure TFrm_principal.Empreendimentos1Click(Sender: TObject);
begin
  if not Verif_senha('Loteamento','Tela de Loteamento','') Then Exit;
//  Frm_Loteamento.Left := Frm_principal.Left+7;
//  Frm_Loteamento.Top := Frm_principal.Top+135;
//  Frm_Loteamento.show;

  try
    if Frm_Loteamento=nil then
       Frm_Loteamento := TFrm_Loteamento.Create(self);
    Frm_Loteamento.Left := Frm_principal.Left+7;
    Frm_Loteamento.Top := Frm_principal.Top+135;
    Frm_Loteamento.showModal;
  finally
    Frm_Loteamento:=nil;
    Frm_Loteamento.Free;
  end;

end;

procedure TFrm_principal.ContaBancria1Click(Sender: TObject);
begin
//  FrmCad_ContasBancarias.Left := Frm_principal.Left+7;
//  FrmCad_ContasBancarias.Top := Frm_principal.Top+135;
//  FrmCad_ContasBancarias.show;
  try
    if FrmCad_ContasBancarias=nil then
       FrmCad_ContasBancarias := TFrmCad_ContasBancarias.Create(self);
    FrmCad_ContasBancarias.Left := Frm_principal.Left+7;
    FrmCad_ContasBancarias.Top := Frm_principal.Top+135;
    FrmCad_ContasBancarias.showModal;
  finally
    FrmCad_ContasBancarias:=nil;
    FrmCad_ContasBancarias.Free;
  end;

end;

procedure TFrm_principal.P1Click(Sender: TObject);
begin
   if not Verif_senha('Estatíticas','Painel','') then exit;
  if frmPainel=nil then
     frmPainel:=TfrmPainel.create(self);
  frmPainel.showmodal;
  FreeAndNil(frmPainel);
end;

procedure TFrm_principal.PlanosdeContas1Click(Sender: TObject);
begin
  if not Verif_senha('Plano de Contas','Tela Principal','') then exit;
  //FrmCad_PlanodeContas.show;
  try
    if FrmCad_PlanodeContas=nil then
       FrmCad_PlanodeContas := TFrmCad_PlanodeContas.Create(self);
    FrmCad_PlanodeContas.Left := Frm_principal.Left+7;
    FrmCad_PlanodeContas.Top := Frm_principal.Top+135;
    FrmCad_PlanodeContas.showModal;
  finally
    FrmCad_PlanodeContas:=nil;
    FrmCad_PlanodeContas.Free;
  end;

end;

procedure TFrm_principal.Cidades1Click(Sender: TObject);
begin
  if not Verif_senha('Cidade','Tela de Cidade','') Then Exit;
//  Frm_Cidade.Left := Frm_principal.Left+7;
//  Frm_Cidade.Top := Frm_principal.Top+135;
//  Frm_Cidade.show;
  try
    if Frm_Cidade=nil then
       Frm_Cidade := TFrm_Cidade.Create(self);
    Frm_Cidade.Left := Frm_principal.Left+7;
    Frm_Cidade.Top := Frm_principal.Top+135;
    Frm_Cidade.showModal;
  finally
    Frm_Cidade:=nil;
    Frm_Cidade.Free;
  end;

end;

procedure TFrm_principal.Empresas1Click(Sender: TObject);
begin
 if not Verif_senha('Empresa','Tela Principal','') then exit;
 try
    if FrmCad_Empresa=nil then
       FrmCad_Empresa := TFrmCad_Empresa.Create(self);
    FrmCad_Empresa.Left := Frm_principal.Left+7;
    FrmCad_Empresa.Top := Frm_principal.Top+135;
    FrmCad_Empresa.showModal;
 finally
    FrmCad_Empresa:=nil;
    FrmCad_Empresa.Free;
 end;
//  FrmCad_Empresa.Left := Frm_principal.Left+7;
//  FrmCad_Empresa.Top := Frm_principal.Top+135;
//  FrmCad_Empresa.show;

end;

procedure TFrm_principal.Configuraes1Click(Sender: TObject);
begin
  if not Verif_senha('Configurações','Tela Principal','') then exit;
  Frm_Configuracoes.Left := Frm_principal.Left+7;
  Frm_Configuracoes.Top := Frm_principal.Top+135;
  Frm_Configuracoes.show;
end;

procedure TFrm_principal.AcertaVendaemRecebimento1Click(Sender: TObject);
begin
//  Frm_acertavenda_em_recebimento.showmodal;
 try
    if Frm_acertavenda_em_recebimento=nil then
       Frm_acertavenda_em_recebimento := TFrm_acertavenda_em_recebimento.Create(self);
    Frm_acertavenda_em_recebimento.Left := Frm_principal.Left+7;
    Frm_acertavenda_em_recebimento.Top := Frm_principal.Top+135;
    Frm_acertavenda_em_recebimento.showModal;
  finally
    Frm_acertavenda_em_recebimento:=nil;
    Frm_acertavenda_em_recebimento.Free;
  end;

end;

procedure TFrm_principal.Fechar1Click(Sender: TObject);
begin
  Close;
end;

procedure TFrm_principal.InibirExibirBarradeBotes1Click(Sender: TObject);
var
  barra:TIniFile;
  VarPath,VarExibe:String;
begin                          
  VarPath := ExtractFilePath( Application.ExeName );
  barra := tIniFile.Create(varpath+'barra.Ini');
  VarExibe := barra.ReadString('BARRA_DE_BOTOES','EXIBE', VarExibe );
  if toolbar.Visible=true then
  begin
     barra.WriteString('BARRA_DE_BOTOES','EXIBE','N');
     VarExibe:='S';
     toolbar.Visible:=false;
  end
  else
  begin
     toolbar.Visible:=true;
     barra.WriteString('BARRA_DE_BOTOES','EXIBE','S');
  end;
  barra.Free;
end;

procedure TFrm_principal.ContasReceberPagar1Click(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Tela Principal','') then exit;
//  FrmCad_Recebimento.Left := Frm_principal.Left+7;
//  FrmCad_Recebimento.Top := Frm_principal.Top+135;
//  FrmCad_Recebimento.show;
  Panel1.Visible:=true;
  Application.ProcessMessages;

  if FrmCad_Recebimento=nil then
     FrmCad_Recebimento:=TFrmCad_Recebimento.Create(Application);
  FrmCad_Recebimento.Show;

end;

procedure TFrm_principal.ReajustedeParcelas1Click(Sender: TObject);
begin
//  Frm_ReajusteDeParcelas.showmodal;
  try
    if Frm_ReajusteDeParcelas=nil then
       Frm_ReajusteDeParcelas := TFrm_ReajusteDeParcelas.Create(self);
    Frm_ReajusteDeParcelas.Left := Frm_principal.Left+7;
    Frm_ReajusteDeParcelas.Top := Frm_principal.Top+135;
    Frm_ReajusteDeParcelas.showModal;
  finally
    Frm_ReajusteDeParcelas:=nil;
    Frm_ReajusteDeParcelas.Free;
  end;

end;

procedure TFrm_principal.Aditamento1Click(Sender: TObject);
begin
  if not Verif_senha('Aditamento','Tela Principal','') then exit;
  if FrmAditamento=nil then
     FrmAditamento:=TFrmAditamento.Create(Self);
  FrmAditamento.show;
end;

procedure TFrm_principal.Quitao1Click(Sender: TObject);
begin
  if not Verif_senha('Quitação','Tela Principal','') then exit;
  if Frmquitacao=nil then
     Frmquitacao:=TFrmquitacao.Create(Self);
  Frmquitacao.show;
end;

procedure TFrm_principal.CessoTransferncia1Click(Sender: TObject);
begin
  if not Verif_senha('Cessao','Tela Principal','') then exit;
  if Frm_Cessao=nil then
     Frm_Cessao:=TFrm_Cessao.Create(Self);
  Frm_Cessao.show;
end;

procedure TFrm_principal.Financeiro1Click(Sender: TObject);
begin
  if not Verif_senha('Financeiro','Tela Principal','') then exit;
//  FrmCad_Recebimento.Left := Frm_principal.Left+7;
//  FrmCad_Recebimento.Top := Frm_principal.Top+135;
//  FrmCad_Recebimento.show;
  Panel1.Visible:=true;
  Application.ProcessMessages;


  DM_tabelas.ZQTipodoc.close;
  DM_tabelas.ZQTipodoc.SQL.Clear;
  DM_tabelas.ZQTipodoc.sql.Add('select * from tipodocumento');
  DM_tabelas.ZQTipodoc.open;
  DM_tabelas.ZQTipodoc.First;

  DM_tabelas.ZQPlanoDeContas.close;
  DM_tabelas.ZQPlanoDeContas.SQL.clear;
  DM_tabelas.ZQPlanoDeContas.SQL.Add('Select * from Plano_contas');
  DM_tabelas.ZQPlanoDeContas.open;

  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select * from RecBxHist limit 100');
  DM_tabelas.ZQReBxHi.open;


  DM_TAbelas.ZQRecebBxTemp.Close;
  DM_TAbelas.ZQRecebBxTemp.SQL.Clear;
  DM_TAbelas.ZQRecebBxTemp.SQL.Add('SELECT * FROM (recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib)join receb_baixa as B ON B.refbaixa=H.refer limit 100');
  DM_TAbelas.ZQRecebBxTemp.open;

  DM_Tabelas.ZQCheque.Close;
  DM_Tabelas.ZQCheque.SQL.Clear;
  DM_Tabelas.ZQCheque.SQL.Add('select * from cheque_rec');
  DM_Tabelas.ZQCheque.open;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento limit 100');
  DM_tabelas.ZQRecebimento.Open;
  DM_Tabelas.ZQRecebimento.Last;

  DM_Tabelas.ZQContaBancaria.SQL.Clear;
  DM_Tabelas.ZQContaBancaria.SQL.Add('Select * from conta_bancaria');
  DM_Tabelas.ZQContaBancaria.Open;



  if FrmCad_Recebimento=nil then
     FrmCad_Recebimento:=TFrmCad_Recebimento.Create(Application);
  FrmCad_Recebimento.ZQGerou.Open;
  FrmCad_Recebimento.Show;
  Panel1.Visible:=false;
  Application.ProcessMessages;

end;

procedure TFrm_principal.abelaPrice1Click(Sender: TObject);
begin
  if not Verif_senha('Tabela Price','Tela da Tabela Price','') Then Exit;
//  Frm_Price.Left := Frm_principal.Left+7;
//  Frm_Price.Top := Frm_principal.Top+135;
//  Frm_Price.show;
  try
    if Frm_Price=nil then
       Frm_Price := TFrm_Price.Create(self);
    Frm_Price.Left := Frm_principal.Left+7;
    Frm_Price.Top := Frm_principal.Top+135;
    Frm_Price.showModal;
  finally
    Frm_Price:=nil;
    Frm_Price.Free;
  end;
end;

procedure TFrm_principal.dxButton2Click(Sender: TObject);
begin
  if not Verif_senha('Tabela Price','Tela da Tabela Price','') Then Exit;
//  Frm_Price.Left := Frm_principal.Left+7;
//  Frm_Price.Top := Frm_principal.Top+135;
//  Frm_Price.show;
  try
    if Frm_Price=nil then
       Frm_Price := TFrm_Price.Create(self);
    Frm_Price.Left := Frm_principal.Left+7;
    Frm_Price.Top := Frm_principal.Top+135;
    Frm_Price.showModal;
  finally
    Frm_Price:=nil;
    Frm_Price.Free;
  end;

end;

procedure TFrm_principal.AcertaCdigodeOrigememRecebimento1Click(
  Sender: TObject);
begin
//  Frm_acertaorigem_em_recebimento.showmodal;
 try
    if Frm_acertaorigem_em_recebimento=nil then
       Frm_acertaorigem_em_recebimento := TFrm_acertaorigem_em_recebimento.Create(self);
    Frm_acertaorigem_em_recebimento.Left := Frm_principal.Left+7;
    Frm_acertaorigem_em_recebimento.Top := Frm_principal.Top+135;
    Frm_acertaorigem_em_recebimento.showModal;
  finally
    Frm_acertaorigem_em_recebimento:=nil;
    Frm_acertaorigem_em_recebimento.Free;
  end;
end;


procedure TFrm_principal.AcertaPrticipanteXRecebimento1Click(
  Sender: TObject);
begin
//  Frm_acertaadversa_recebimento.showmodal;
 try
    if Frm_acertaadversa_recebimento=nil then
       Frm_acertaadversa_recebimento := TFrm_acertaadversa_recebimento.Create(self);
    Frm_acertaadversa_recebimento.Left := Frm_principal.Left+7;
    Frm_acertaadversa_recebimento.Top := Frm_principal.Top+135;
    Frm_acertaadversa_recebimento.showModal;
  finally
    Frm_acertaadversa_recebimento:=nil;
    Frm_acertaadversa_recebimento.Free;
  end;

end;

procedure TFrm_principal.A1Click(Sender: TObject);
begin
  if not Verif_senha('Utilitários','Acerto dos valores das Parcelas ','') then exit;
//  Frm_Acerto_parcelas.showmodal;
  try
    if Frm_Acerto_parcelas=nil then
       Frm_Acerto_parcelas := TFrm_Acerto_parcelas.Create(self);
    Frm_Acerto_parcelas.Left := Frm_principal.Left+7;
    Frm_Acerto_parcelas.Top := Frm_principal.Top+135;
    Frm_Acerto_parcelas.showModal;
  finally
    Frm_Acerto_parcelas:=nil;
    Frm_Acerto_parcelas.Free;
  end;


end;

procedure TFrm_principal.A2Click(Sender: TObject);
begin
  if not Verif_senha('Alteração da Data de Vencimento','Alterar Vencimento Parcelas','') Then Exit;
//  Frm_ProximoVctoParcelas.showmodal;
  try
    if Frm_ProximoVctoParcelas=nil then
       Frm_ProximoVctoParcelas := TFrm_ProximoVctoParcelas.Create(self);
    Frm_ProximoVctoParcelas.Left := Frm_principal.Left+7;
    Frm_ProximoVctoParcelas.Top := Frm_principal.Top+135;
    Frm_ProximoVctoParcelas.showModal;
  finally
    Frm_ProximoVctoParcelas:=nil;
    Frm_ProximoVctoParcelas.Free;
  end;


end;

procedure TFrm_principal.A3Click(Sender: TObject);
begin
  try
    if FrmAgenda = nil then
       FrmAgenda:=TFrmAgenda.create(self);
    FrmAgenda.showmodal;
  finally
    FreeAndNil(FrmAgenda);
  end;
end;

procedure TFrm_principal.abelaIGPM1Click(Sender: TObject);
begin
  if not Verif_senha('Tabela IGPM','Tela da Tabela IGPM','') Then Exit;
//  Frm_igpm.Left := Frm_principal.Left+7;
//  Frm_igpm.Top := Frm_principal.Top+135;
//  Frm_igpm.show;
  try
    if Frm_igpm=nil then
       Frm_igpm := TFrm_igpm.Create(self);
    Frm_igpm.Left := Frm_principal.Left+7;
    Frm_igpm.Top := Frm_principal.Top+135;
    Frm_igpm.showModal;
  finally
    Frm_igpm:=nil;
    Frm_igpm.Free;
  end;

end;

procedure TFrm_principal.AcertaDatadoPrximoReajuste1Click(Sender: TObject);
begin
//    Frm_acertaproximo_reajuste_em_recebimento.showmodal;
  try
    if Frm_acertaproximo_reajuste_em_recebimento=nil then
       Frm_acertaproximo_reajuste_em_recebimento := TFrm_acertaproximo_reajuste_em_recebimento.Create(self);
    Frm_acertaproximo_reajuste_em_recebimento.Left := Frm_principal.Left+7;
    Frm_acertaproximo_reajuste_em_recebimento.Top := Frm_principal.Top+135;
    Frm_acertaproximo_reajuste_em_recebimento.showModal;
  finally
    Frm_acertaproximo_reajuste_em_recebimento:=nil;
    Frm_acertaproximo_reajuste_em_recebimento.Free;
  end;

end;

procedure TFrm_principal.iposdeDocumentos1Click(Sender: TObject);
begin
  if not Verif_senha('Tipo de Documentos','Tela Principal','') then exit;
  try
    if FrmCad_TipoDoc=nil then
       FrmCad_TipoDoc := TFrmCad_TipoDoc.Create(self);
    FrmCad_TipoDoc.Left := Frm_principal.Left+7;
    FrmCad_TipoDoc.Top := Frm_principal.Top+135;
    FrmCad_TipoDoc.showModal;
  finally
    FrmCad_TipoDoc:=nil;
    FrmCad_TipoDoc.Free;
  end;
//  FrmCad_TipoDoc.show;
end;

procedure TFrm_principal.AcertaroValorFinaldaVenda1Click(Sender: TObject);
begin
//  Frm_acertavalor_final_venda.showmodal;
  try
    if Frm_acertavalor_final_venda=nil then
       Frm_acertavalor_final_venda := TFrm_acertavalor_final_venda.Create(self);
    Frm_acertavalor_final_venda.Left := Frm_principal.Left+7;
    Frm_acertavalor_final_venda.Top := Frm_principal.Top+135;
    Frm_acertavalor_final_venda.showModal;
  finally
    Frm_acertavalor_final_venda:=nil;
    Frm_acertavalor_final_venda.Free;
  end;

end;

procedure TFrm_principal.ExcluirLoteamento1Click(Sender: TObject);
begin
  try
    if Frm_Excluir=nil then
       Frm_Excluir := TFrm_Excluir.Create(self);
    Frm_Excluir.Left := Frm_principal.Left+7;
    Frm_Excluir.Top := Frm_principal.Top+135;
    Frm_Excluir.showModal;
  finally
    Frm_Excluir:=nil;
    Frm_Excluir.Free;
  end;

//  Frm_Excluir.showmodal;
end;

procedure TFrm_principal.ColocardataeHoradoAditamentonamesmaLinha1Click(
  Sender: TObject);
begin
//  Frm_acertaaditamento.showmodal;
  try
    if Frm_acertaaditamento=nil then
       Frm_acertaaditamento := TFrm_acertaaditamento.Create(self);
    Frm_acertaaditamento.Left := Frm_principal.Left+7;
    Frm_acertaaditamento.Top := Frm_principal.Top+135;
    Frm_acertaaditamento.showModal;
  finally
    Frm_acertaaditamento:=nil;
    Frm_acertaaditamento.Free;
  end;

end;

procedure TFrm_principal.AlteraroCampoPrximoReajuste1Click(
  Sender: TObject);
begin
  if not Verif_senha('Alteração da Data do Próximo Reajuste','Alterar Próximo Reajuste','') Then Exit;
//  Frm_ProximoReajusteDeParcelas.showmodal;
  try
    if Frm_ProximoReajusteDeParcelas=nil then
       Frm_ProximoReajusteDeParcelas := TFrm_ProximoReajusteDeParcelas.Create(self);
    Frm_ProximoReajusteDeParcelas.Left := Frm_principal.Left+7;
    Frm_ProximoReajusteDeParcelas.Top := Frm_principal.Top+135;
    Frm_ProximoReajusteDeParcelas.showModal;
  finally
    Frm_ProximoReajusteDeParcelas:=nil;
    Frm_ProximoReajusteDeParcelas.Free;
  end;

end;

procedure TFrm_principal.JBarraUtilItems3Click(Sender: TObject);
begin
  if not Verif_senha('Utilitários','Mês/Ano Reajuste','') then exit;
  JBarraUtil.Visible:=false;
//  Frm_Acerto_ano.showmodal;
  try
    if Frm_Acerto_ano=nil then
       Frm_Acerto_ano := TFrm_Acerto_ano.Create(self);
    Frm_Acerto_ano.Left := Frm_principal.Left+7;
    Frm_Acerto_ano.Top := Frm_principal.Top+135;
    Frm_Acerto_ano.showModal;
  finally
    Frm_Acerto_ano:=nil;
    Frm_Acerto_ano.Free;
  end;

end;

procedure TFrm_principal.JBarraUtilItems4Click(Sender: TObject);
begin
  JBarraUtil.Visible := False;
 try
    if Frm_Zerar_nosso_numero=nil then
       Frm_Zerar_nosso_numero   := TFrm_Zerar_nosso_numero.Create(self);
    Frm_Zerar_nosso_numero.Left := Frm_principal.Left+7;
    Frm_Zerar_nosso_numero.Top  := Frm_principal.Top+135;
    Frm_Zerar_nosso_numero.showModal;
  finally
    Frm_Zerar_nosso_numero:=nil;
    Frm_Zerar_nosso_numero.Free;
  end;
end;

procedure TFrm_principal.AdicionarMsAnodeReajuste1Click(Sender: TObject);
begin
  if not Verif_senha('Utilitários','Mês/Ano Reajuste','') then exit;
//  Frm_Acerto_ano.showmodal;
  try
    if Frm_Acerto_ano=nil then
       Frm_Acerto_ano := TFrm_Acerto_ano.Create(self);
    Frm_Acerto_ano.Left := Frm_principal.Left+7;
    Frm_Acerto_ano.Top := Frm_principal.Top+135;
    Frm_Acerto_ano.showModal;
  finally
    Frm_Acerto_ano:=nil;
    Frm_Acerto_ano.Free;
  end;

end;

procedure TFrm_principal.Email1Click(Sender: TObject);
begin
   if not Verif_senha('Participante','Tela Email','') Then Exit;
   ExecutePrograma (ExtractFilePath(Application.ExeName)+'\email.exe', '');
end;

procedure TFrm_principal.Cadastro1Click(Sender: TObject);
begin
  if not Verif_senha('Participante','Tela Principal','') Then Exit;
//  Frm_participante.Left := Frm_principal.Left+7;
//  Frm_participante.Top := Frm_principal.Top+135;
//  Frm_Participante.show;
  try
    if Frm_Participante=nil then
       Frm_Participante := TFrm_Participante.Create(self);
//    Frm_Participante.Left := Frm_principal.Left+7;
//    Frm_Participante.Top := Frm_principal.Top+135;
   // Frm_Participante.showModal;
   Frm_Participante.show;
  finally
//    Frm_Participante:=nil;
//    Frm_Participante.Free;
  end;

end;

procedure TFrm_principal.FormCreate(Sender: TObject);
begin
{  if DM_Tabelas=nil then
     DM_Tabelas:=tDM_Tabelas.Create(application);
  DM_Tabelas.ZQConfiguracoes.Open;
  //chama resolução de tela
   PTela(Sender);
 }
  ConfigurarMenuModerno;
end;

procedure TFrm_principal.FormDestroy(Sender: TObject);
begin
  DM_Tabelas:=nil;
  //volta a resolução antiga
  PTela(Sender);

end;

procedure TFrm_principal.ConsultaRpida1Click(Sender: TObject);
begin
  if not Verif_senha('Participante','Consulta Rápida','') Then Exit;
//  frmparticipante2.Left := Frm_principal.Left+7;
//  frmparticipante2.Top := Frm_principal.Top+135;
//  frmparticipante2.show;
  try
    if FrmParticipante2=nil then
       FrmParticipante2 := TFrmParticipante2.Create(self);
    FrmParticipante2.Left := Frm_principal.Left+7;
    FrmParticipante2.Top := Frm_principal.Top+135;
    FrmParticipante2.showModal;
  finally
    FrmParticipante2:=nil;
    FrmParticipante2.Free;
  end;

end;

procedure TFrm_principal.LimparTabelaRemessaRecebidaporPerodo1Click(
  Sender: TObject);
begin
    if not Verif_senha('Utilitarios','Apagar Remessa Recebida por Período','') Then Exit;
    try
      if frm_acerta_remessa=nil then
         frm_acerta_remessa := Tfrm_acerta_remessa.Create(self);
      frm_acerta_remessa.Left := Frm_principal.Left+7;
      frm_acerta_remessa.Top := Frm_principal.Top+135;
      frm_acerta_remessa.showModal;
    finally
      frm_acerta_remessa:=nil;
      frm_acerta_remessa.Free;
    end;

//    frm_acerta_remessa.showmodal;
end;


procedure TFrm_principal.dxButton3Click(Sender: TObject);
begin
  if not Verif_senha('Participante','Consulta Rápida','') Then Exit;
//  frmparticipante2.Left := Frm_principal.Left+7;
//  frmparticipante2.Top := Frm_principal.Top+135;
//  frmparticipante2.show;
  try
    if FrmParticipante2=nil then
       FrmParticipante2 := TFrmParticipante2.Create(self);
//    FrmParticipante2.Left := Frm_principal.Left+7;
//    FrmParticipante2.Top := Frm_principal.Top+135;
//    FrmParticipante2.showModal;
      frmparticipante2.show;
  finally
//    FrmParticipante2:=nil;
//    FrmParticipante2.Free;
  end;

end;

procedure TFrm_principal.dxButton19Click(Sender: TObject);
begin
  try
    if FrmCad_ContasBancarias=nil then
       FrmCad_ContasBancarias := TFrmCad_ContasBancarias.Create(self);
    FrmCad_ContasBancarias.Left := Frm_principal.Left+7;
    FrmCad_ContasBancarias.Top := Frm_principal.Top+135;
    FrmCad_ContasBancarias.showModal;
  finally
    FrmCad_ContasBancarias:=nil;
    FrmCad_ContasBancarias.Free;
  end;

end;

procedure TFrm_principal.JXBarraCad_DiversosItems0Click(Sender: TObject);
begin
   JXBarraCad_Diversos.Visible := False;
  if not Verif_senha('Tipo de Documentos','Tela Principal','') then exit;
//  FrmCad_TipoDoc.show;
  try
    if FrmCad_TipoDoc=nil then
       FrmCad_TipoDoc := TFrmCad_TipoDoc.Create(self);
    FrmCad_TipoDoc.Left := Frm_principal.Left+7;
    FrmCad_TipoDoc.Top := Frm_principal.Top+135;
    FrmCad_TipoDoc.showModal;
  finally
    FrmCad_TipoDoc:=nil;
    FrmCad_TipoDoc.Free;
  end;


end;

procedure TFrm_principal.JBarraUtilItems0Click(Sender: TObject);
begin
  JBarraUtil.Visible := False;
end;

procedure TFrm_principal.JBarraUtilItems1Click(Sender: TObject);
begin
  JBarraUtil.Visible := False;
end;

procedure TFrm_principal.Button1Click(Sender: TObject);
var
  I: Integer;
  FrmAlerta: TForm;
  BarraIniciar: HWND; {Barra Iniciar}
  tmAltura: Integer;
  tmRect: TRect;
  xTop : Integer;
  xIniTop : Integer;

begin
  FrmAlerta := TFrmAlerta.Create(self);
  //localiza o Handle da janela iniciar
  BarraIniciar := FindWindow('Shell_TrayWnd', nil);
  //Pega o "retângulo" que envolve a barra e sua altura
  GetWindowRect(BarraIniciar, tmRect);
  tmAltura := tmRect.Bottom - tmRect.Top;

  with FrmAlerta do
  begin
    Left := Screen.Width - ClientWidth;
    if tmRect.Top = -2 then
      tmAltura := 30;
    //Pega o top final
    xTop := Screen.Height - ClientHeight - tmAltura;
    //Pega o top inicial
    xIniTop := Screen.Height + ClientHeight + tmAltura;
    Top := xIniTop;
    for I := xIniTop downto xTop do
    begin
      Top := Top - 1;
      Show;
      Update;
      Application.ProcessMessages;
      Sleep(10);
    end;
    Sleep(10000);
    Close;
    FreeAndNil(FrmAlerta);
  end;

end;

procedure TFrm_principal.Timer2Timer(Sender: TObject);
 var
 MainHandle : THandle;

begin
 // liberar memoria
 try
   MainHandle := OpenProcess(PROCESS_ALL_ACCESS, false, GetCurrentProcessID) ;
   SetProcessWorkingSetSize(MainHandle, $FFFFFFFF, $FFFFFFFF) ;
   CloseHandle(MainHandle) ;
 except
 end;
 Application.ProcessMessages;
 StatusBar1.Panels[1].Text := FormatFloat('Mem. Usada: ,.# K', CurrentMemoryUsage / 1024);
end;

procedure TFrm_principal.Timer3Timer(Sender: TObject);
var
  barra:TIniFile;
  VarPath,VarExibe:String;
begin
  timer3.Enabled:=false;
  VarPath := ExtractFilePath( Application.ExeName );
  barra := tIniFile.Create(varpath+'barra.Ini');
  try
    VarExibe := barra.ReadString('BARRA_DE_BOTOES','EXIBE', VarExibe );
    if empty(VarExibe) then
    Begin
       barra.WriteString('BARRA_DE_BOTOES','EXIBE','S');
       VarExibe:='S';
    end;
    if VarExibe='N' then
    begin
       toolbar.Visible:=false;
    end
    else
    begin
      toolbar.Visible:=true;
    end;
  finally
    barra.Free;
  end;
end;

procedure TFrm_principal.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  { ZQErro_Baixa executa UPDATE e nao possui campos de resultado. }

//  RLConsts.SetVersion(3,71,'B');


end.

