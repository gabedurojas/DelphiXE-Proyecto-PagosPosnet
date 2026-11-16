unit Menu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,  ADODB, StrUtils,
  jpeg, pngimage, ComCtrls, JvExExtCtrls, JvExtComponent, JvPanel,
  ExtCtrls, JvExStdCtrls, JvBehaviorLabel, JvExControls, JvLabel, ShellApi,
  Menus, JvMenus, JvNavigationPane, JvSpeedbar, ToolWin, JvExComCtrls,
  JvToolBar, Buttons, JvExButtons, JvBitBtn, GR32_Image, JvEmbeddedForms,
  RzLabel, RzGroupBar, Data.DB, RzTreeVw, Vcl.ImgList;

type HackImg = class(TImage);
type
  TForm1 = class(TForm)
    Panel3: TPanel;
    JvPanel2: TJvPanel;
    JvBehaviorLabel2: TJvBehaviorLabel;
    Label1: TLabel;
    Panel1: TPanel;
    Image321: TImage32;
    RzLabel1: TRzLabel;
    EmbedPanel: TJvEmbeddedFormPanel;
    MenuAccesos: TRzGroupBar;
    Planes: TRzGroup;
    barra: TStatusBar;
    Contrib: TRzGroup;
    Intima: TRzGroup;
    BoletaDeuda: TRzGroup;
    PPAbogados: TRzGroup;
    ImageList2: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure lblCloseClick(Sender: TObject);
    procedure OP1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Panel3MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Panel1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Image2Click(Sender: TObject);
    procedure EmbedPanelFormDestroy(Sender: TObject);
    procedure Autor1Items0Click(Sender: TObject);
    procedure MenuAccesosCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure Convenio1Items0Click(Sender: TObject);
    procedure RzGroup1Items0Click(Sender: TObject);
    procedure Convenio1Items4Click(Sender: TObject);
    procedure Convenio1Items3Click(Sender: TObject);
    procedure Parametria1Items0Click(Sender: TObject);
    procedure Parametria1Items1Click(Sender: TObject);
    function GetAppVersion: string;
    procedure PlanesItems5Click(Sender: TObject);
    procedure PlanesItems11Click(Sender: TObject);
    procedure PlanesItems12Click(Sender: TObject);
    procedure ContribItems0Click(Sender: TObject);
    procedure ContribItems1Click(Sender: TObject);
    procedure ContribItems2Click(Sender: TObject);
    procedure PlanesItems6Click(Sender: TObject);
    procedure PlanesItems4Click(Sender: TObject);
    procedure PlanesItems9Click(Sender: TObject);
    procedure IntimaItems0Click(Sender: TObject);
    procedure IntimaItems1Click(Sender: TObject);
    procedure IntimaItems2Click(Sender: TObject);
    procedure IntimaItems3Click(Sender: TObject);
    procedure ContribItems3Click(Sender: TObject);
    procedure IntimaItems4Click(Sender: TObject);
    procedure BoletaDeudaItems0Click(Sender: TObject);
    procedure BoletaDeudaItems1Click(Sender: TObject);
    procedure BoletaDeudaItems2Click(Sender: TObject);
    procedure BoletaDeudaItems3Click(Sender: TObject);
    procedure VerArea(nro:byte);
    procedure VerArea2(nro:byte);
    procedure BoletaDeudaItems4Click(Sender: TObject);
    procedure PPAbogadosItems0Click(Sender: TObject);

  private
    gh:boolean;
    dato,dato1:string;
    DataSet2,DataSet1:TADODataSet;
    { Private declarations }
  public

  end;

var
  Form1: TForm1;

implementation

uses Unidad, Acceso1, Modulo, Acerca, A_PlanPago,Con_Planes,
     Imp_CuotasPlanes, PlanPagoModifDom, AM_PlanPagpPagare,
     Anular_PagarePPagos,ConsPagare, Canc_Objetos, AsignRecep_PP, Resumen_PP,
     AM_Contrib, Conf_ContribuDebitos, Report_Debitos,
     ResumenAuditoria_PP,Imp_LosPeriodosdelPlan,Intimaciones_Buscar,
     Intimaciones_ImpNew,Intimaciones_CamcioEstado,Intimaciones_Consultas,
     IncluirAniosPrescriptos,BoletaDeuda_Buscar,BoletaDeuda_ImpNew,
     BoletaDeuda_CambioEstado,BoletoDeuda_Consultas,BoletaDeuda_Bloqueo,
     A_PlanPagoAbogados;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
gh:=false;
DataSet2:=TADODataSet.Create(nil);
DataSet1:=TADODataSet.Create(nil);
Modulo2.Sistema.ExecVentana:=false;
Image321.Visible:=true;
EmbedPanel.Visible:=false;
end;

procedure TForm1.VerArea(nro:byte);
var t,b:integer;
    Cnx2:TADOConnection;
begin
  Cnx2:=TADOConnection.Create(nil);
  Cnx2.LoginPrompt:=false;
  if nro=1 then begin
    Cnx2.ConnectionString:=Modulo2.Sistema.ConectDb;
    DataSet2.Close;
    DataSet2.Connection:=Cnx2;
    DataSet2.CommandText:='SELECT NroUsuario as Nrousuario FROM Fn_PlanPagoUsuariosAdmin  (' + modulo2.Usuario.UsuNro + ')';
    DataSet2.Open;
    if not Dataset2.IsEmpty  then dato:=DataSet2.Fields[0].AsString
      else dato:='0';
    DataSet2.Connection:=nil;
  end;
  Cnx2.Close;
  FreeAndNil(DataSet2);
  FreeAndNil(Cnx2);
end;

procedure TForm1.VerArea2(nro:byte);
var t,b:integer;
    Cnx2:TADOConnection;
begin
  Cnx2:=TADOConnection.Create(nil);
  Cnx2.LoginPrompt:=false;
  if nro=1 then begin
    Cnx2.ConnectionString:=Modulo2.Sistema.ConectInti;
    DataSet1.Close;
    DataSet1.Connection:=Cnx2;
    DataSet1.CommandText:='SELECT NroUsuario as Nrousuario FROM Fn_BoletaDeudaUsuariosAdmin  (' + modulo2.Usuario.UsuNro + ')';
    DataSet1.Open;
    if not Dataset1.IsEmpty  then dato1:=DataSet1.Fields[0].AsString
      else dato1:='0';
    DataSet1.Connection:=nil;
  end;
  Cnx2.Close;
  FreeAndNil(DataSet1);
  FreeAndNil(Cnx2);
end;

procedure TForm1.BoletaDeudaItems0Click(Sender: TObject);
var  BoletaDeuda_Busca: TBoletaDeuda_Busca;
begin
  if Modulo2.ObtenerNivelAcceso('173') > 1 then   begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
   + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
   MenuAccesos.Visible:=false;
   Modulo2.Sistema.ExecVentana:=true;
   Image321.Visible:=false;
   EmbedPanel.Visible:=true;
   EmbedPanel.AlwaysVisible:=true;
   BoletaDeuda_Busca:= TBoletaDeuda_Busca.create(self);
   EmbedPanel.FormLink:=BoletaDeuda_Busca.FormularioEmb;
   BoletaDeuda_Busca.OnActivate(sender);
  end;
end;

procedure TForm1.BoletaDeudaItems1Click(Sender: TObject);
var  BolDeuda_ImpNew: TBolDeuda_ImpNew;
begin
  if Modulo2.ObtenerNivelAcceso('173') > 1 then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    BolDeuda_ImpNew:= TBolDeuda_ImpNew.create(self);
    EmbedPanel.FormLink:=BolDeuda_ImpNew.FormularioEmb;
    BolDeuda_ImpNew.Opcion:=true;
    BolDeuda_ImpNew.OnActivate(sender);
  end;
end;

procedure TForm1.BoletaDeudaItems2Click(Sender: TObject);
var  BoletaDeuda_CambEstado: TBoletaDeuda_CambEstado;
begin
  if Modulo2.ObtenerNivelAcceso('173') > 1 then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    BoletaDeuda_CambEstado:= TBoletaDeuda_CambEstado.create(self);
    EmbedPanel.FormLink:=BoletaDeuda_CambEstado.FormularioEmb;
    BoletaDeuda_CambEstado.OnActivate(sender);
  end;
end;

procedure TForm1.BoletaDeudaItems3Click(Sender: TObject);
var  Consultas_BoletoDeuda: TConsultas_BoletoDeuda;
begin
 // if Modulo2.ObtenerNivelAcceso('173') > 1 then   begin
 //     messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
 //     + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
 // end
 // else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    Consultas_BoletoDeuda:= TConsultas_BoletoDeuda.create(self);
    EmbedPanel.FormLink:=Consultas_BoletoDeuda.FormularioEmb;
    Consultas_BoletoDeuda.OnActivate(sender);
 // end;
end;

procedure TForm1.BoletaDeudaItems4Click(Sender: TObject);
var  BolDeuda_Bloqueo: TBolDeuda_Bloqueo;
begin
  if Modulo2.ObtenerNivelAcceso('173') > 1 then   begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
   + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
   MenuAccesos.Visible:=false;
   Modulo2.Sistema.ExecVentana:=true;
   Image321.Visible:=false;
   EmbedPanel.Visible:=true;
   EmbedPanel.AlwaysVisible:=true;
   BolDeuda_Bloqueo:= TBolDeuda_Bloqueo.create(self);
   EmbedPanel.FormLink:=BolDeuda_Bloqueo.FormularioEmb;
   BolDeuda_Bloqueo.OnActivate(sender);
  end;
end;

procedure TForm1.ContribItems0Click(Sender: TObject);
var AM_Contribuyentes: TAM_Contribuyentes;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
AM_Contribuyentes:= TAM_Contribuyentes.create(self);
AM_Contribuyentes.NroOpcion:=1;
EmbedPanel.FormLink:=AM_Contribuyentes.FormularioEmb;
AM_Contribuyentes.OnActivate(sender);
end;

procedure TForm1.ContribItems1Click(Sender: TObject);
var Cons_Contrib: TCons_Contrib;
begin
 if Modulo2.ObtenerNivelAcceso('151') >=2 then begin
    messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
    + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
 end
 else begin
  MenuAccesos.Visible:=false;
  Modulo2.Sistema.ExecVentana:=true;
  Image321.Visible:=false;
  EmbedPanel.Visible:=true;
  EmbedPanel.AlwaysVisible:=true;
  Cons_Contrib:= TCons_Contrib.create(self);
  EmbedPanel.FormLink:=Cons_Contrib.FormularioEmb;
  Cons_Contrib.OnActivate(sender);
 end;
end;

procedure TForm1.ContribItems2Click(Sender: TObject);
var Consul_Debitos: TConsul_Debitos;
begin
if Modulo2.ObtenerNivelAcceso('151') >= 3 then begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
    + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
end
else begin
  MenuAccesos.Visible:=false;
  Modulo2.Sistema.ExecVentana:=true;
  Image321.Visible:=false;
  EmbedPanel.Visible:=true;
  EmbedPanel.AlwaysVisible:=true;
  Consul_Debitos:= TConsul_Debitos.create(self);
  EmbedPanel.FormLink:=Consul_Debitos.FormularioEmb;
  Consul_Debitos.OnActivate(sender);
end;
end;

procedure TForm1.ContribItems3Click(Sender: TObject);
var AM_Contribuyentes: TAM_Contribuyentes;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
AM_Contribuyentes:= TAM_Contribuyentes.create(self);
AM_Contribuyentes.NroOpcion:=2;
EmbedPanel.FormLink:=AM_Contribuyentes.FormularioEmb;
AM_Contribuyentes.OnActivate(sender);
end;

procedure TForm1.Convenio1Items0Click(Sender: TObject);
var Alta_PlanPago: TAlta_PlanPago;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
Alta_PlanPago:= TAlta_PlanPago.create(self);
EmbedPanel.FormLink:=Alta_PlanPago.FormularioEmb;
Alta_PlanPago.OnActivate(sender);
end;

function TForm1.GetAppVersion: string;
var
    Size, Size2: DWord;
    Pt, Pt2: Pointer;
begin
  Size := GetFileVersionInfoSize(PChar(ParamStr(0)), Size2);
  if Size > 0 then
  begin
    GetMem(Pt, Size);
    try
     GetFileVersionInfo(PChar(ParamStr(0)), 0, Size, Pt);
     VerQueryValue(Pt, '\', Pt2, Size2);
     with TVSFixedFileInfo(Pt2^) do begin
      Result := '' +
      IntToStr(HiWord(dwFileVersionMS)) +  IntToStr(LoWord(dwFileVersionMS)) +
      IntToStr(HiWord(dwFileVersionLS)) +  IntToStr(LoWord(dwFileVersionLS)) ;
      //IntToStr(HiWord(dwFileVersionMS)) + '.' +
      //IntToStr(LoWord(dwFileVersionMS)) + '.' +
      //IntToStr(HiWord(dwFileVersionLS));
      //+ ' Build ' + IntToStr(LoWord(dwFileVersionLS));
     end;
   finally
      FreeMem(Pt);
   end;
  end;
end;

procedure TForm1.Convenio1Items3Click(Sender: TObject);
var PPModifDom: TPPModifDom;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
PPModifDom:= TPPModifDom.create(self);
EmbedPanel.FormLink:=PPModifDom.FormularioEmb;
PPModifDom.OnActivate(sender);
end;

procedure TForm1.Convenio1Items4Click(Sender: TObject);
var ImpresionPlanes: TImpresionPlanes;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
ImpresionPlanes:= TImpresionPlanes.create(self);
EmbedPanel.FormLink:=ImpresionPlanes.FormularioEmb;
ImpresionPlanes.OnActivate(sender);
end;

procedure TForm1.EmbedPanelFormDestroy(Sender: TObject);
begin
Image321.Visible:=true;
EmbedPanel.Visible:=false;
Modulo2.Sistema.ExecVentana:=false;
end;

procedure TForm1.FormActivate(Sender: TObject);
var Acceso:TAcceso;
    path:string;
begin
 if not(gh) then begin
  path:=GetCurrentDir() + '\Imagenes\LOGO08.JPG';
  if FileExists(path) then  Image321.Bitmap.LoadFromFile(path);
  Acceso:=TAcceso.create(self);
  try
    Acceso.ShowModal;
  finally
    if Acceso.control = false then close;
    gh:=true;
    Acceso.destroy;
    MenuAccesos.Enabled:=true;
  end;
 end;
 barra.Panels[0].Text:='Nro Usuario: ' + Modulo2.Usuario.UsuNro;
 barra.Panels[1].Text:='Nombre Usuario: ' + Modulo2.Usuario.UsuNombre;
 barra.Panels[2].Text:='Version - ' + GetAppVersion;
 VerArea(1);
 if dato <> '0' then planes.Items[6].Visible:=true
    else planes.Items[6].Visible:=false;

 if Modulo2.ObtenerNivelAcceso('152') <> 0 then begin
  Planes.Opened:=true;
  Planes.Enabled:=true;
 end
 else begin
  Planes.Opened:=false;
  Planes.Enabled:=false;
 end;

 if Modulo2.ObtenerNivelAcceso('151') <> 0 then begin
  Contrib.Opened:=true;
  Contrib.Enabled:=true;
 end
 else begin
  Contrib.Opened:=false;
  Contrib.Enabled:=false;
 end;

 if Modulo2.ObtenerNivelAcceso('162') <> 0 then begin
  intima.Opened:=true;
  intima.Enabled:=true;
 end
 else begin
  intima.Opened:=false;
  intima.Enabled:=false;
 end;

 if Modulo2.ObtenerNivelAcceso('173') <> 0 then begin
  BoletaDeuda.Opened:=true;
  BoletaDeuda.Enabled:=true;
  VerArea2(1);
  if dato1 <> '0' then BoletaDeuda.Items[4].Visible:=true
    else BoletaDeuda.Items[4].Visible:=false;
 end
 else begin
  BoletaDeuda.Opened:=false;
  BoletaDeuda.Enabled:=false;
 end;

 if Modulo2.ObtenerNivelAcceso('205') <> 0 then begin
  PPAbogados.Opened:=true;
  PPAbogados.Enabled:=true;
 end
 else begin
  PPAbogados.Opened:=false;
  PPAbogados.Enabled:=false;
 end;

end;

procedure TForm1.lblCloseClick(Sender: TObject);
begin
Close;
end;

procedure TForm1.OP1Click(Sender: TObject);
begin
Self.Height:=659;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action:=caFree;
end;

procedure TForm1.Panel3MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
ReleaseCapture();
Perform( WM_SYSCOMMAND, $F012, 0 );
end;

procedure TForm1.Parametria1Items0Click(Sender: TObject);
 var AM_Pagares: TAM_Pagares;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
AM_Pagares:= TAM_Pagares.create(self);
EmbedPanel.FormLink:=AM_Pagares.FormularioEmb;
AM_Pagares.OnActivate(sender);
end;

procedure TForm1.RzGroup1Items0Click(Sender: TObject);
var  ConsultaPlanes: TConsultaPlanes;
begin
  if Modulo2.ObtenerNivelAcceso('152') <> 1  then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    ConsultaPlanes:= TConsultaPlanes.create(self);
    EmbedPanel.FormLink:=ConsultaPlanes.FormularioEmb;
    ConsultaPlanes.OnActivate(sender);
  end;
end;

procedure TForm1.MenuAccesosCustomDrawItem(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
with Sender as TRzTreeView do begin
  if Node.Level = 0 then begin
    Canvas.Font.Name:='MS Sans Serif';
    Canvas.Font.Color:=clWindowText;
    Canvas.Font.Style:=Canvas.Font.Style + [fsBold];
    Canvas.Font.Size:=11;
  end
  else begin
    Canvas.Font.Name:='Tahoma';
    Canvas.Font.Color:=clNavy;
    Canvas.Font.Size:=9;
  end;
end;
end;

procedure TForm1.Autor1Items0Click(Sender: TObject);
var Acerca_de:TAcerca_de;
begin
try
  MenuAccesos.Visible:=false;
  Acerca_de:=TAcerca_de.Create(Nil);
  Acerca_de.ShowModal;
finally
  FreeAndNil(Acerca_de);
end;
end;

procedure TForm1.Panel1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if not Modulo2.Sistema.ExecVentana then MenuAccesos.Visible:=true;
MenuAccesos.Color:=$00FECBEC;
end;

procedure TForm1.Image2Click(Sender: TObject);
var Acceso:TAcceso;
begin
try
  Acceso:=TAcceso.Create(Nil);
  Acceso.ShowModal;
  MenuAccesos.Enabled:=Modulo2.ConAcceso;
finally
  FreeAndNil(Acceso);
end;
end;

procedure TForm1.IntimaItems0Click(Sender: TObject);
var  Intima_Busca: TIntima_Busca;
begin
  if Modulo2.ObtenerNivelAcceso('162') > 1 then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    Intima_Busca:= TIntima_Busca.create(self);
    EmbedPanel.FormLink:=Intima_Busca.FormularioEmb;
    Intima_Busca.OnActivate(sender);
  end;
end;

procedure TForm1.IntimaItems1Click(Sender: TObject);
var  Intima_Imp: TIntima_ImpNew;
begin
  if Modulo2.ObtenerNivelAcceso('162') > 3 then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    Intima_Imp:= TIntima_ImpNew.create(self);
    EmbedPanel.FormLink:=Intima_Imp.FormularioEmb;
    Intima_Imp.OnActivate(sender);
  end;
end;

procedure TForm1.IntimaItems2Click(Sender: TObject);
var  Intima_CambiaEstado: TIntima_CambiaEstado;
begin
  if Modulo2.ObtenerNivelAcceso('162') > 1 then   begin
    messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
    + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
   MenuAccesos.Visible:=false;
   Modulo2.Sistema.ExecVentana:=true;
   Image321.Visible:=false;
   EmbedPanel.Visible:=true;
   EmbedPanel.AlwaysVisible:=true;
   Intima_CambiaEstado:= TIntima_CambiaEstado.create(self);
   EmbedPanel.FormLink:=Intima_CambiaEstado.FormularioEmb;
   Intima_CambiaEstado.OnActivate(sender);
  end;
end;

procedure TForm1.IntimaItems3Click(Sender: TObject);
var  Consultas_Intimaciones: TConsultas_Intimaciones;
begin
  if Modulo2.ObtenerNivelAcceso('162') > 3 then   begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
   + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
   MenuAccesos.Visible:=false;
   Modulo2.Sistema.ExecVentana:=true;
   Image321.Visible:=false;
   EmbedPanel.Visible:=true;
   EmbedPanel.AlwaysVisible:=true;
   Consultas_Intimaciones:= TConsultas_Intimaciones.create(self);
   EmbedPanel.FormLink:=Consultas_Intimaciones.FormularioEmb;
   Consultas_Intimaciones.OnActivate(sender);
  end;
end;

procedure TForm1.IntimaItems4Click(Sender: TObject);
var  AgregaAniosaPlanes: TAgregaAniosaPlanes;
begin
  if Modulo2.ObtenerNivelAcceso('162') > 1 then   begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
   + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    AgregaAniosaPlanes:= TAgregaAniosaPlanes.create(self);
    EmbedPanel.FormLink:=AgregaAniosaPlanes.FormularioEmb;
    AgregaAniosaPlanes.OnActivate(sender);
  end;
end;

procedure TForm1.Parametria1Items1Click(Sender: TObject);
var  ConsultaPagare: TConsultaPagare;
begin
  if Modulo2.ObtenerNivelAcceso('152') <> 1 then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    ConsultaPagare:= TConsultaPagare.create(self);
    EmbedPanel.FormLink:=ConsultaPagare.FormularioEmb;
    ConsultaPagare.OnActivate(sender);
  end;
end;

procedure TForm1.PlanesItems11Click(Sender: TObject);
var  ResumenPP: TResumenPP;
begin
  if Modulo2.ObtenerNivelAcceso('152') < 1  then begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    ResumenPP:= TResumenPP.create(self);
    ResumenPP.NroOpcion:=1;
    EmbedPanel.FormLink:=ResumenPP.FormularioEmb;
    ResumenPP.OnActivate(sender);
  end;
end;

procedure TForm1.PlanesItems12Click(Sender: TObject);
var  AsignaRecepcion: TAsignaRecepcion;
begin
  if Modulo2.ObtenerNivelAcceso('152') <> 1  then begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    AsignaRecepcion:= TAsignaRecepcion.create(self);
    EmbedPanel.FormLink:=AsignaRecepcion.FormularioEmb;
    AsignaRecepcion.OnActivate(sender);
  end;
end;

procedure TForm1.PlanesItems4Click(Sender: TObject);
var ImpresionPeriodosPlanes: TImpresionPeriodosPlanes;
begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
ImpresionPeriodosPlanes:= TImpresionPeriodosPlanes.create(self);
EmbedPanel.FormLink:=ImpresionPeriodosPlanes.FormularioEmb;
ImpresionPeriodosPlanes.OnActivate(sender);
end;

procedure TForm1.PlanesItems5Click(Sender: TObject);
var  CancelaObjetos: TCancelaObjetos;
begin
  if Modulo2.ObtenerNivelAcceso('152') <> 1  then   begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    CancelaObjetos:= TCancelaObjetos.create(self);
    EmbedPanel.FormLink:=CancelaObjetos.FormularioEmb;
    CancelaObjetos.OnActivate(sender);
  end;
end;

procedure TForm1.PlanesItems6Click(Sender: TObject);
 var AuditoriaPP: TAuditoriaPP;
begin
if Modulo2.ObtenerNivelAcceso('152') <> 1 then begin
   messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
    + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
end
else begin
MenuAccesos.Visible:=false;
Modulo2.Sistema.ExecVentana:=true;
Image321.Visible:=false;
EmbedPanel.Visible:=true;
EmbedPanel.AlwaysVisible:=true;
AuditoriaPP:= TAuditoriaPP.create(self);
EmbedPanel.FormLink:=AuditoriaPP.FormularioEmb;
AuditoriaPP.OnActivate(sender);
end;
end;

procedure TForm1.PlanesItems9Click(Sender: TObject);
var  ResumenPP: TResumenPP;
begin
  if Modulo2.ObtenerNivelAcceso('152') < 1  then begin
      messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
      + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
  end
  else begin
    MenuAccesos.Visible:=false;
    Modulo2.Sistema.ExecVentana:=true;
    Image321.Visible:=false;
    EmbedPanel.Visible:=true;
    EmbedPanel.AlwaysVisible:=true;
    ResumenPP:= TResumenPP.create(self);
    ResumenPP.NroOpcion:=2;
    EmbedPanel.FormLink:=ResumenPP.FormularioEmb;
    ResumenPP.OnActivate(sender);
  end;
end;

procedure TForm1.PPAbogadosItems0Click(Sender: TObject);
var Alta_PlanPagoAbogados: TAlta_PlanPagoAbogados;
begin
 if Modulo2.ObtenerNivelAcceso('205') = 1  then begin
   MenuAccesos.Visible:=false;
  Modulo2.Sistema.ExecVentana:=true;
  Image321.Visible:=false;
  EmbedPanel.Visible:=true;
  EmbedPanel.AlwaysVisible:=true;
  Alta_PlanPagoAbogados:= TAlta_PlanPagoAbogados.create(self);
  EmbedPanel.FormLink:=Alta_PlanPagoAbogados.FormularioEmb;
  Alta_PlanPagoAbogados.OnActivate(sender);
 end
 else begin
  messagedlg('ATENCION.. Esta ventana NO esta Habilitada' + chr(13)
  + 'para el Usuario Nro: ' + Modulo2.Usuario.UsuNro + ' Consulte con un Administrador...',mtInformation,[mbok],0);
 end;
end;

end.
