unit PagosPosnet;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Mask,
  Vcl.ComCtrls, JvgPage, Vcl.ImgList, JvComponentBase, JvEmbeddedForms, Data.DB,
  Vcl.Grids, Vcl.DBGrids, JvExDBGrids, JvDBGrid,StrUtils,ADODB, JvExButtons,
  JvBitBtn, Vcl.ExtCtrls, JvExMask, JvToolEdit, JvMaskEdit, JvCheckedMaskEdit,
  JvDatePickerEdit;

  type THackStringGrid = class(TStringGrid);
type
    TLosPagosPosnet = class(TForm)
    PC1: TJvgPageControl;
    TabSheet1: TTabSheet;
    GroupBox4: TGroupBox;
    FormularioEmb: TJvEmbeddedFormLink;
    ImageList1: TImageList;
    DataSource1: TDataSource;
    GroupBox3: TGroupBox;
    sale: TBitBtn;
    grupoperiodos: TGroupBox;
    ASIENTOPERIODO: TComboBox;
    buscar: TBitBtn;
    Label11: TLabel;
    L2: TEdit;
    Label14: TLabel;
    Edit6: TEdit;
    Label5: TLabel;
    ordenPago: TEdit;
    Label27: TLabel;
    L1: TEdit;
    Grouporecibo: TGroupBox;
    BitBtn1: TBitBtn;
    Edit2: TEdit;
    Label2: TLabel;
    JvDBGrid2: TJvDBGrid;
    Edit1: TEdit;
    Label3: TLabel;
    JvDBGrid1: TJvDBGrid;
    Label1: TLabel;
    Label6: TLabel;
    Edit3: TEdit;
    DataSource2: TDataSource;
    GroupBox1: TGroupBox;
    INSERTAR: TBitBtn;
    ok: TImage;
    NO: TImage;
    Label7: TLabel;
    grillaidentidad: TStringGrid;
    Label8: TLabel;
    Label4: TLabel;
    ComboTarjeta: TComboBox;
    Edit4: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CargaCombos(nro:byte);
    procedure Consultas(nro:byte);
    procedure Limpiar(nro:byte);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ASIENTOPERIODOExit(Sender: TObject);
    procedure ASIENTOPERIODOKeyPress(Sender: TObject; var Key: Char);
    procedure JvDBGrid2CellClick(Column: TColumn);
    procedure INSERTARClick(Sender: TObject);
    procedure L1Exit(Sender: TObject);
    procedure L2Exit(Sender: TObject);
    procedure L2KeyPress(Sender: TObject; var Key: Char);
    procedure saleClick(Sender: TObject);
    procedure Edit6KeyPress(Sender: TObject; var Key: Char);
    procedure ordenPagoKeyPress(Sender: TObject; var Key: Char);
    procedure Edit6Exit(Sender: TObject);
    procedure ordenPagoExit(Sender: TObject);
    procedure Edit2Exit(Sender: TObject);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure buscarClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure JvDBGrid1CellClick(Column: TColumn);
    procedure grillaidentidadClick(Sender: TObject);
    procedure ordenPagoChange(Sender: TObject);
    procedure ComboTarjetaExit(Sender: TObject);
    procedure ComboTarjetaKeyPress(Sender: TObject; var Key: Char);

  private
   IdCombo:array of array of string;
   gh,modif,det,elim:boolean;
   FechaAct,elidasiento,txi,elidasientodet,feccierre,detasiento,idtarje:string;
   Com1,Com2,Com3:TADOCommand;
   Cnx2:TADOConnection;
   DataSet1,ASIENTOS,ASIENTOSDET,DataSet4,DataSetAlta,MAXIMOASIENTO,DataSet3:TADODataSet;
   totd,toth:real;
   g1,g2:Extended ;
    { Private declarations }
  public
   control:boolean;
    { Public declarations }
  end;

var
  LosPagosPosnet: TLosPagosPosnet;

implementation

{$R *.dfm}

uses Unidad, modulo;

procedure TLosPagosPosnet.FormClose(Sender: TObject; var Action: TCloseAction);
begin
freeAndNil(ASIENTOS);
FREEANDNIL(ASIENTOSDET);
FREEANDNIL(DataSet4);
FreeAndNil(DataSet3);
FreeAndNil(DataSet1);
end;

procedure TLosPagosPosnet.FormCreate(Sender: TObject);
begin
  gh:=false;
  ASIENTOS:=TADODataSet.Create(Self);
  ASIENTOSDET:=TADODataSet.Create(Self);
  DataSet4:=TADODataSet.Create(Self);
  DataSet3:=TADODataSet.Create(nil);
   DataSet1:=TADODataSet.Create(nil);
end;

procedure TLosPagosPosnet.grillaidentidadClick(Sender: TObject);
begin
 if not ASIENTOSDET.IsEmpty then begin
  insertar.Enabled:=true;
  insertar.SetFocus;
 end;
end;

procedure TLosPagosPosnet.FormActivate(Sender: TObject);
begin
 if not (gh) then begin
  control:=false;
  grillaidentidad.Cells[0,0]:='idrecibo';              grillaidentidad.Cells[1,0]:='IdLiquidacion';
  grillaidentidad.Cells[2,0]:='IdTasaAnual';           grillaidentidad.Cells[3,0]:='fechaemision';
  grillaidentidad.Cells[4,0]:='IdCaja';                grillaidentidad.Cells[5,0]:='IdCajero';
  grillaidentidad.Cells[6,0]:='importe';               grillaidentidad.Cells[7,0]:='anio';
  grillaidentidad.Cells[8,0]:='puntov';

  SetLength(IdCombo,1);
  gh:=true;
  feccierre:='';
  FechaAct:=LeftStr(Fec_System(),10);
  PC1.ActivePage:=PC1.Pages[0];
  modif:=false;
  det:=false;
  elim:=false;
  grupoperiodos.Visible :=true;
  //Consultas(0);
  grupoperiodos.SetFocus;
  //CargaCombos(1);
  CargaCombos(2);
  //ASIENTOPERIODO.ItemIndex:=0;
  combotarjeta.ItemIndex :=0;
  combotarjeta.SetFocus;
  edit3.Visible:=false;
  ok.Visible:=false;
  NO.Visible:=false;
  insertar.Enabled:=false;
  buscar.Enabled:=false;
  label7.Caption:='...';
 end;
end;

procedure TLosPagosPosnet.JvDBGrid1CellClick(Column: TColumn);
begin
 if not ASIENTOSDET.IsEmpty then begin
  insertar.Enabled:=true;
  insertar.SetFocus;
 end;
end;

procedure TLosPagosPosnet.JvDBGrid2CellClick(Column: TColumn);
begin
 if not ASIENTOS.IsEmpty then begin
  edit1.Text:=ASIENTOS.Fields[0].AsString;
  Consultas(2);
 end;
end;

procedure TLosPagosPosnet.ASIENTOPERIODOExit(Sender: TObject);
begin
 if ASIENTOPERIODO.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una Teminal válida',mterror,[mbok],0);
  ASIENTOPERIODO.SetFocus;
 end;
end;

procedure TLosPagosPosnet.ASIENTOPERIODOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  L2.SetFocus;
 end;
end;

procedure TLosPagosPosnet.BitBtn1Click(Sender: TObject);
begin
Consultas(1);
end;

procedure TLosPagosPosnet.buscarClick(Sender: TObject);
var gf:integer;
    txtw:string;
var Cnx2:TADOConnection;
begin
  gf:=messagedlg('Confirma GUARDAR los datos del Posnet?',mtWarning,[mbok,mbCancel],0);
  if gf = 1 then begin
   try
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.ConectDbMuni;
    DataSet3:=TADODataSet.Create(nil);
    DataSet3.Connection:=Cnx2;
    DataSet3.CommandText:='exec PRC_PosnetPagos_Insertar ' + trim(asientoperiodo.text)
    + ' , ' + L2.Text
    + ' , ' + edit6.Text
    + ' , ' + ordenpago.Text
    + ' , ' + edit4.Text;
    DataSet3.Open;
    txi:=DataSet3.Fields[0].AsString;
    DataSet3.Close;
    DataSet3.Connection:=nil;
    messagedlg('El Nuevo Dato del Posnet fue Guardado correctamente..',mtConfirmation,[mbok],0);
    l1.text:=  txi;
    if l1.text <> '' then begin
     Grouporecibo.Enabled:=true;
     edit2.setfocus;
    end
    else begin
     Grouporecibo.Enabled:=false;
     buscar.setfocus;
    end;
   except
    EnviarLog(self.Caption,DataSetAlta.CommandText);
    messagedlg('ERROR al Insertar el Pago Posnet.. Comuniquese con un Administrador...',mterror,[mbok],0);
   end;
   Cnx2.Close;
   FreeAndNil(Cnx2);
  end;
end;

procedure TLosPagosPosnet.INSERTARClick(Sender: TObject);
var gf,a:integer;
    txtw:string;
var Cnx2:TADOConnection;
begin
  gf:=messagedlg('Confirma GUARDAR los datos del PosnetPagoRecibo?',mtWarning,[mbok,mbCancel],0);
  if gf = 1 then begin
   try
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.ConectDbMuni;
    DataSet4:=TADODataSet.Create(nil);
    DataSet4.Connection:=Cnx2;
    Com1:=TADOCommand.Create(nil);
    Com1.Connection:=Cnx2;

    for gf:=1 to grillaidentidad.RowCount-1 do begin
     Com1.CommandText:='PRC_PosnetPagosRecibos_Insertar  ' + L1.Text
     + ' , ' + chr(39) + trim(grillaidentidad.Cells[0,gf]) + chr(39);
     Com1.Execute;
    end;

    ok.Visible:=true;
    no.Visible:=false;
    label7.Caption:='El PosnetPagoRecibo fue Guardado correctamente..';

    Limpiar(1);
    pc1.Pages[0].TabVisible:=true;
    PC1.ActivePage:=PC1.Pages[0];
   except
    EnviarLog(self.Caption,Com1.CommandText);
    ok.Visible:=false;
    no.Visible:=true;
    label7.Caption:='ERROR al Guardar el PosnetPagoRecibo.. ';
   end;
   FreeAndNil(Com1);
   Cnx2.Close;
   FreeAndNil(Cnx2);
  end;
end;

procedure TLosPagosPosnet.CargaCombos(nro:byte);
begin
  if nro=1 then Modulo2.CargaCombo(ASIENTOPERIODO,'SELECT IdTerminal FROM Posnet_Terminales WHERE (Deshabilitado = 0) order by 1 asc',Modulo2.ConectDbMuni);
  if nro=2 then Modulo2.CargaCombo(ComboTarjeta,'SELECT Descripcion FROM Posnet_Tarjetas where Deshabilitada = 0 ORDER BY IdTarjeta',Modulo2.ConectDbMuni);
end;

procedure TLosPagosPosnet.ComboTarjetaExit(Sender: TObject);
begin
 if ComboTarjeta.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Tipo de Tarjeta válida',mterror,[mbok],0);
  ComboTarjeta.SetFocus;
 end
 else Consultas(0);
end;

procedure TLosPagosPosnet.ComboTarjetaKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  CargaCombos(1);
  asientoperiodo.SetFocus;
 end;
end;

procedure TLosPagosPosnet.Consultas(nro:byte);
var t,y,gf,p,u:integer;
begin
 Cnx2:=TADOConnection.Create(nil);
 Cnx2.LoginPrompt:=false;
 Cnx2.ConnectionString:=modulo2.ConectDbMuni;
 if nro = 0 then begin
  DataSet1.Close;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SELECT IdTarjeta  FROM [municipio].[dbo].[Posnet_Tarjetas]  '
  + ' where [Descripcion] = ' + chr(39) + trim(combotarjeta.text) + chr(39) + ' and Deshabilitada = 0 ';
  DataSet1.Open;
  if not DataSet1.IsEmpty then begin
    idtarje:= trim(DataSet1.Fields[0].AsString);
    edit4.Text := idtarje;
  end;
  DataSet1.Connection:=nil;
 end;
 if nro = 1 then begin
  ASIENTOS.Close;
  ASIENTOS.Connection:=Cnx2;
  DataSource1.DataSet:=ASIENTOS;
  ASIENTOS.CommandText:='SELECT idregistro,cuim,partida,nro_doc,nomraz,monto1v,'
  +' primerv,codbar43,puntov,idrubro,codigoobjeto,SubRubro FROM tasaanual WHERE (codbar43 = ' + trim(edit2.text) + ')';
  ASIENTOS.Open;
  if not ASIENTOS.IsEmpty  then begin
   JvDBGrid2.Columns[0].Width := 90;
   JvDBGrid2.Columns[1].Width := 90;
   JvDBGrid2.Columns[2].Width := 90;
   JvDBGrid2.Columns[3].Width := 90;
   JvDBGrid2.Columns[4].Width := 250;
   JvDBGrid2.Columns[5].Width := 90;
   JvDBGrid2.Columns[6].Width := 90;
   JvDBGrid2.Columns[7].Width := 90;
   JvDBGrid2.Columns[8].Width := 90;
  end;
  ASIENTOS.Connection:=nil;
  Cnx2.Close;
 end;
 if nro = 2 then begin
  ASIENTOSDET.Close;
  ASIENTOSDET.Connection:=Cnx2;
  DataSource2.DataSet:=ASIENTOSDET;
  ASIENTOSDET.CommandText:='SELECT idrecibo,IdLiquidacion,IdTasaAnual,fechaemision,'
  +' IdCaja,IdCajero,importe,anio,puntov FROM TasaAnualRecibos WHERE (IdTasaAnual =' + trim(edit1.text) + ')';
  ASIENTOSDET.Open;
  for t:=1 to ASIENTOSDET.RecordCount do begin
   grillaidentidad.RowCount:=t+1;
   ASIENTOSDET.RecNo:=t;
   grillaidentidad.Cells[0,t]:=ASIENTOSDET.Fields[0].AsString;
   grillaidentidad.Cells[1,t]:=ASIENTOSDET.Fields[1].AsString;
   grillaidentidad.Cells[2,t]:=ASIENTOSDET.Fields[2].AsString;
   grillaidentidad.Cells[3,t]:=ASIENTOSDET.Fields[3].AsString;
   grillaidentidad.Cells[4,t]:=ASIENTOSDET.Fields[4].AsString;
   grillaidentidad.Cells[5,t]:=ASIENTOSDET.Fields[5].AsString;
   grillaidentidad.Cells[6,t]:=ASIENTOSDET.Fields[6].AsString;
   grillaidentidad.Cells[7,t]:=ASIENTOSDET.Fields[7].AsString;
   grillaidentidad.Cells[8,t]:=ASIENTOSDET.Fields[8].AsString;
  end;
  if not ASIENTOSDET.IsEmpty  then insertar.visible:=true
    else  insertar.visible:=false;
  //end;
  ASIENTOSDET.Connection:=nil;
  Cnx2.Close;
 { for p:=0 to grillacuentas.ColCount-1 do
    for u:=1 to grillacuentas.RowCount-1 do grillacuentas.Cells[p,u]:='';
  grillacuentas.RowCount:=2;
  totd:=0;
  toth:=0;
  ASIENTOSDET.Close;
  ASIENTOSDET.Connection:=Cnx2;
  ASIENTOSDET.CommandText:='select IdAsiento,IdCuentaFinanciera,CuentaFinanciera,DetalleCuentaFinanciera,'
  +'Correlativo,Descripcion,Debe,Haber,'
  +'IdDetalle,Columna FROM Fn_AsientosContablesDetalle ( ' + trim(elidasiento) +') order by Correlativo';
  ASIENTOSDET.Open;
  for t:=1 to ASIENTOSDET.RecordCount do begin
   grillacuentas.RowCount:=t+1;
   ASIENTOSDET.RecNo:=t;
   grillacuentas.Cells[0,t]:=ASIENTOSDET.Fields[0].AsString;   //IdAsiento
   grillacuentas.Cells[1,t]:=ASIENTOSDET.Fields[1].AsString;   //IdCuentaFinanciera
   grillacuentas.Cells[2,t]:=ASIENTOSDET.Fields[2].AsString;   //CuentaFinanciera
   grillacuentas.Cells[3,t]:=ASIENTOSDET.Fields[3].AsString;   //DetalleCuentaFinanciera
   grillacuentas.Cells[4,t]:=ASIENTOSDET.Fields[4].AsString;   //Correlativo
   grillacuentas.Cells[5,t]:=ASIENTOSDET.Fields[5].AsString;   //Descripcion
   grillacuentas.Cells[6,t]:=ASIENTOSDET.Fields[6].AsString;   //Debe
   grillacuentas.Cells[7,t]:=ASIENTOSDET.Fields[7].AsString;   //haber
   grillacuentas.Cells[8,t]:=ASIENTOSDET.Fields[8].AsString;   //IdDetalle
   grillacuentas.Cells[9,t]:=ASIENTOSDET.Fields[9].AsString;   //Columna
   grillacuentas.Cells[10,t]:='';  //Modif
   grillacuentas.Cells[11,t]:='';  //select
  end;

  for y:=1 to grillacuentas.RowCount-1 do begin
   if (grillacuentas.Cells[6,y] <> '') and (grillacuentas.Cells[12,y] = '') then begin
    totd:=totd + strtofloat(grillacuentas.Cells[6,y]) ;
   end;
  end;

  for y:=1 to grillacuentas.RowCount-1 do begin
   if (grillacuentas.Cells[7,y] <> '') and (grillacuentas.Cells[12,y] = '') then begin
    toth:=toth + strtofloat(grillacuentas.Cells[7,y]) ;
   end;
  end;
 
  ASIENTOSDET.Connection:=nil;  }
 end;
end;

procedure TLosPagosPosnet.Edit2Exit(Sender: TObject);
begin
 if (trim(Edit2.Text) = '') OR (trim(Edit2.Text) = '0') then begin
  messagedlg('Error!! No Existe un Nº de Recibo ingresado.',mterror,[mbok],0);
  Edit2.SetFocus;
 end;
end;

procedure TLosPagosPosnet.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  bitbtn1.SetFocus;
 end
 else begin
  if (StrScan('0123456789',Key) <> nil) or (Key = Char(VK_BACK)) then
   else Key := #0;
 end;
end;

procedure TLosPagosPosnet.Edit6Exit(Sender: TObject);
begin
 if (trim(Edit6.Text) = '') OR (trim(Edit6.Text) = '0')  then begin
  messagedlg('Error!! No Existe un Nº ticket ingresado.',mterror,[mbok],0);
  Edit6.SetFocus;
 end;
end;

procedure TLosPagosPosnet.Edit6KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  ordenPago.SetFocus;
 end
 else begin
  if (StrScan('0123456789',Key) <> nil) or (Key = Char(VK_BACK)) then
   else Key := #0;
 end;
end;

procedure TLosPagosPosnet.Limpiar(nro:byte);
var t,y,gf,p,u:integer;
    totd,toth:real;
begin
 if nro = 1 then begin
  insertar.Enabled :=false;
  edit1.Clear;
  edit3.Clear;
  edit4.Clear;
  l1.Clear;
  ordenpago.Clear;
  edit6.Clear;
  l2.Clear;
  edit1.Clear;
  edit2.Clear;

  for p:=0 to grillaidentidad.ColCount-1 do
    for u:=1 to grillaidentidad.RowCount-1 do grillaidentidad.Cells[p,u]:='';
  grillaidentidad.RowCount:=2;

  ok.Visible:=false;
  no.Visible:=false;
  PC1.ActivePage:=PC1.Pages[0];
  label7.Caption:='...';
  combotarjeta.SetFocus;
 end;
end;

procedure TLosPagosPosnet.L1Exit(Sender: TObject);
begin
 if trim(L1.Text) = '' then begin
  messagedlg('Error!! No Existe Nro. de Asiento Contable ingresado.',mterror,[mbok],0);
  L1.SetFocus;
 end;
end;

procedure TLosPagosPosnet.L2Exit(Sender: TObject);
begin
 if (trim(L2.Text) = '') OR (trim(L2.Text) = '0') then begin
  messagedlg('Error!! No Existe una Descripcion ingresado.',mterror,[mbok],0);
  L2.SetFocus;
 end;
end;

procedure TLosPagosPosnet.L2KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  EDIT6.SetFocus;
 end
 else begin
  if (StrScan('0123456789',Key) <> nil) or (Key = Char(VK_BACK)) then
   else Key := #0;
 end;
end;

procedure TLosPagosPosnet.ordenPagoChange(Sender: TObject);
begin
   if trim(ordenpago.Text) <> '' then BUSCAR.Enabled:=true
     else BUSCAR.Enabled:=false;
end;

procedure TLosPagosPosnet.ordenPagoExit(Sender: TObject);
begin
 if (trim(ordenPago.Text) = '') OR (trim(ordenPago.Text) = '0') then begin
  messagedlg('Error!! No Existe un Importe ingresado.',mterror,[mbok],0);
  ordenPago.SetFocus;
 end;
end;

procedure TLosPagosPosnet.ordenPagoKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  BUSCAR.SetFocus;
 end
 else begin
  if (StrScan('0123456789.',Key) <> nil) or (Key = Char(VK_BACK)) then
   else Key := #0;
 end;
end;

procedure TLosPagosPosnet.saleClick(Sender: TObject);
begin
 close;
end;

end.
