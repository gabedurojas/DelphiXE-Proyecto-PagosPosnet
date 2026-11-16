unit Objetos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes;


type TSistema = Class(TObject)
  Public
    ConectDb:widestring;
    ConectUsu:widestring;
    ConectSeg:widestring;
    ConectInti:widestring;
    ConectBolDeuda:widestring;
    MyHost:string;
    MyIP:string;
    Server:string;
    Path_todo:string;
    ByteTransf:integer;
    ExecVentana:boolean;
    NroAplic:string;
    LogComentario:string;
    NombreAplic:string;
    ConAcceso:boolean;
    Camara:string;
    Esquema:string;
    Comando:string;
    Fondo:string;
    PuntoV:string;
   constructor Create();
 end;

Type
 TComboClase =Class(TObject)
 Public
    Param1:String;
    Param2:String;
    Param3:String;
    Param4:String;

   constructor Create();
end;

type TUsuario = Class(TObject)
  Public
    UsuNro:string;
    UsuNroArea:TStringList;
    UsuRed:string;
    UsuNvl:TStringList;
    UsuMail:string;
    UsuNombre:string;
    UsuDetalleNivel:TStringList;
    UsuPuntoV:String;
    Clave:string;
    Intentos:integer;
    Vencimiento:string;

   constructor Create();
 end;

type TObjetos = Class(TObject)
  Public
    CantObjetos:integer;
    Identidad:TStringList;
    Rubro:TStringList;
    EsContenedor:TStringList;
    ConDeuda:TStringList;
    CodObjeto:TStringList;
    Descripcion:TStringList;
    EsPlan:TStringList;
    ImportePagado:TStringList;
    DetalleImpu:TStringList;
    BajaProvisoria :TStringList;
    constructor Create();
 end;

type TRecibo = Class(TObject)
  Public
    NombreRespo:string;
    CuimRespo:string;
    DomicilioRespo:string;
    NombreTit:string;
    CuimTit:string;
    DomicilioTit:string;
    CuotaDe:string;
    Anio:string;
    NroPlan:string;
    NroCuota:string;
    FechaIni:string;
    CantObj:string;
    Objetos:widestring;
    DetalleRecibo:string;
    NroCobro:string;
    NroLiquidacion:string;
    PrimerV:string;
    Importe1:string;
    SegundoV:string;
    Importe2:string;
    CodBar:string;
    NroCodBar:string;
    NroUsuario:string;
    TituloDesc:string;
    IdRegistro:string;
    SubRubro:string;
    constructor Create();
 end;

 type TFiltro = Class(TObject)
  Public
    Columna:TStringList;
    Tipo:TStringList;
    Valor:string;
    Patron:boolean;

   constructor Create();
 end;

implementation


constructor TSistema.Create();
begin
ConectDb:='';
ConectUsu:='';
ConectBolDeuda:='';
//ConectInti:='';
MyHost:='';
MyIP:='';
Server:='';
Path_todo:='';
ByteTransf:=0;
ExecVentana:=false;
NroAplic:='';
LogComentario:='';
NombreAplic:='';
ConAcceso:=false;
Camara:='';
Esquema:='';
Comando:='';
Fondo:='';
PuntoV:='';
end;

constructor TComboClase.Create();
begin
//Inicio los valores por defecto
Param1:='';
Param2:='';
Param3:='';
Param4:='';
end;

constructor TUsuario.Create();
begin
//Inicio los valores por defecto
UsuNro:='';
UsuNroArea:=TStringList.Create;
UsuRed:='';
UsuNvl:=TStringList.Create;
UsuDetalleNivel:=TStringList.Create;
UsuMail:='';
UsuNombre:='';
Clave:='';
Intentos:=0;
Vencimiento:='';
end;

constructor TObjetos.Create();
begin
CantObjetos:=0;
Identidad:=TStringList.Create;
Rubro:=TStringList.Create;
EsContenedor:=TStringList.Create;
ConDeuda:=TStringList.Create;
CodObjeto:=TStringList.Create;
Descripcion:=TStringList.Create;
EsPlan:=TStringList.Create;
ImportePagado:=TStringList.Create;
DetalleImpu:=TStringList.Create;
BajaProvisoria:=TStringList.Create;
end;

constructor TRecibo.Create();
begin
NombreRespo:='';
CuimRespo:='';
DomicilioRespo:='';
NombreTit:='';
CuimTit:='';
DomicilioTit:='';
CuotaDe:='';
Anio:='';
NroPlan:='';
NroCuota:='';
FechaIni:='';
CantObj:='';
Objetos:='';
DetalleRecibo:='';
NroCobro:='';
NroLiquidacion:='';
PrimerV:='';
Importe1:='';
SegundoV:='';
Importe2:='';
CodBar:='';
NroCodBar:='';
NroUsuario:='';
TituloDesc:='';
IdRegistro:='';
SubRubro:='';
end;

constructor TFiltro.Create();
begin
Columna:=TStringList.Create;
Tipo:=TStringList.Create;
Valor:='';
Patron:=false;
end;

end.
