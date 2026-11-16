unit Unidad;

 interface
uses  SysUtils, DB, ADODB, DateUtils, Dialogs,DBWeb, Classes, Grids, DBGrids,
JvDBGrid;

type lineas = array[1..2] of string;
type StrArray = array [1..9] of string;

//function modificar(text:string):string;
Function edad(fact:string;fnac:string):string;
Function tmpbaja(fecini:string;meses:string):string;
Function fechalarga(fecha:string):string;
function separar(cadena:string;separador:string;posi:integer):string;
Function StrFill(cadena:string;longitud:Integer;relleno,posicion:string): string;
Function cortastr(cadena:string;tamano:integer):lineas;
Function strcut(cadena:string;tamano:integer):string;
Function cuenta(cadena:widestring;separador:string):integer;
Function tomadato(cadena:string;lugar:integer):integer;
function Extrae(cadena:widestring;separador:char;pos:integer):string;
function BuscaChr(cadena:widestring;caracter:string;posi:integer):integer;
function separar1(cadena:string;separador:string):string;
function incorpora(cadena:string;tamano,decimal:integer):string;
function mda(fecha:string;tipo:byte):integer;
function crea_cuim(nro_dni,prefijo:integer):string;
Function llenastr(cadena:string;tamano:integer;caracter:char):string;
Function Cambia(cadena:string;caracter1,caracter2:char):string;
function inttocuim(nro_cuim:integer):string;
function cuimtoint(nro_cuim:string):string;
function inttodni(nro_cuim:string):string;
Procedure AuditAccesos(AreaIngreso:string;OperacionRealizada:char);
Function AuditaImpresos(OperacionRealizada,ElCuim:string):string;
Function fechacompleta(fecha:string):string;
function strtocuim(nro_cuim:string):string;
Function Fec_System():string;
Function Corta_fec(cadena:string):string;
Function V_Modulos(cadena:string):boolean;
function calcular_deuda(fec_ant,fec_act:string):real;
function decimales(cadena:string;lugares:byte):real;
function redondeo(cadena:string;decimales:integer):string;
Function strcut_midle(cadena:widestring;comienzo,fin:integer):string;
Function completa1(cadena:string;tamano:integer;caracter:char):string;
function importetobanco(cadena:string;tamano,decimales:integer;caracter:char):string;
function fecha_dif(f_desde,f_hasta:string):string;
function digito(cadena:string;opc:byte):string;
function codigo_barra(cadena:string):string;
function findchr(cadena:string;caracter:char):integer;
Function Fec_Texto(cadena:string):string;
function el_mes(mes:integer):string;
Function fechaCuotas(cadena:string;cuotas:integer):string;
function fectoint(cadena:string):string;
Function Potencia(x,n:Real):Real;
function numtopesos(num:string):string;
Function completa2(cadena:string;tamano1,tamano2:integer;caracter:char):string;
function transforma(cadena:string;carac1,carac2:char):string;
Function MesIncremento(Fecha:string;Meses:integer):string;
Function DiaIncremento(Fecha:string;Dias:integer):string;
Function AnioIncremento(Fecha:string;Anios:integer):string;
Procedure EnviarLog(NombreVentana:string;Texto:widestring);
function DiasDiferencia(FechaInico,FechaFin:string):integer;
function encripta(cadena:string;tpo:boolean):string;
procedure EncontrarHijos(Codigo,CantTabs,PosRec:integer;Tabla:TADOTable);
Procedure CrearTreeConTbs(FileName,NombreTabla:string;Tabla:TADOTable);
function usu_level(usu,area:integer):integer;
Function CreaPeriodosStrGrid(var Grilla:TStringGrid; ColAnio,ColCuota,ColControl:byte;CharControl:string):widestring;
Function CreaPeriodosDBGrid(var Grilla:TADODataSet; ColAnio,ColCuota,ColControl:byte;CharControl:string):widestring;
function codigo_ver(cadena:string;EsCodbar:boolean):string;
FUNCTION emailValido(CONST Value: String): boolean;
function FormatoMoneda(Valor:string):string;

var ListaOrganigrama:TStringList;
var Contador:array of integer;
    ComUni:TADOCommand;
    DtsUni:TADODataSet;
implementation

uses modulo;

function FormatoMoneda(Valor:string):string;
var dato : currency;
    ver,ver2,ger :string;
begin
  dato := strtocurr(Valor);
  ger := currtostrf(dato,ffCurrency, 2);
  ver:=StringReplace(ger,'.',';',[rfReplaceAll]);
  ver2:= StringReplace(ver,',','.',[rfReplaceAll]);//ver;
  Result:=StringReplace(ver2,';',',',[rfReplaceAll]);
end;


{function modificar(text:string):string;
    var o,u:integer;dato:string[5];
    begin
      o:=1;
      for u:=1 to 5 do begin
        if not((text[u]='_') or (text[u]=' ')) then begin
           dato[o]:=text[u];o:=o+1;
        end;
      end;
      dato[0]:=chr(o-1);
      modificar:=dato;
end;   }

{Funcion para tratamiento de fechas, devuelve la edad}
Function edad(fact:string;fnac:string):string;
var dac,mac,aac:integer;
    dnc,mnc,anc:integer;
    ano,mes,dia:integer;
begin
if fnac = '' then edad:='S/D'
else begin
 dac:=strtoint(fact[1]+fact[2]);
 mac:=strtoint(fact[4]+fact[5]);
 if length(fact)>8 then aac:=strtoint(fact[7]+fact[8]+fact[9]+fact[10])
 else aac:=strtoint('20'+fact[7]+fact[8]);
 dnc:=strtoint(fnac[1]+fnac[2]);
 mnc:=strtoint(fnac[4]+fnac[5]);
 if length(fnac)>8 then anc:=strtoint(fnac[7]+fnac[8]+fnac[9]+fnac[10])
 else anc:=strtoint('19'+fnac[7]+fnac[8]);
 ano:=aac-anc;mes:=mac-mnc;dia:=dac-dnc;
 if mes > 0 then edad:=inttostr(ano)
 else begin
   if mes < 0 then edad:=inttostr(ano-1)
   else begin
      if dia < 0 then edad:=inttostr(ano-1)
      else edad:=inttostr(ano);
   end;
 end;
end;
end;



Function tmpbaja(fecini:string;meses:string):string;
var dnc,mnc,ano:integer;
    aux,mes:integer;
    df,mf,af:string;
begin
 dnc:=strtoint(fecini[1]+fecini[2]);
 mnc:=strtoint(fecini[4]+fecini[5]);
 if length(fecini)>8 then ano:=strtoint(fecini[7]+fecini[8]+fecini[9]+fecini[10])
 else ano:=strtoint('20'+fecini[7]+fecini[8]);
 mes:=mnc + strtoint(trim(meses));
 if mes > 12 then begin
   aux:=mes mod 12;
   mes:=trunc(mes/12);
   ano:=ano + mes;
   mes:=aux;
 end;
 df:=inttostr(dnc);
 if length(df) = 1 then df:='0' + df;
 mf:=inttostr(mes);
 if length(mf) = 1 then mf:='0' + mf;
 af:=inttostr(ano);
 tmpbaja:= df + '/' + mf + '/' + af;
end;

Function fechalarga(fecha:string):string;
var dia,mes,ano:string;
begin
if length(fecha)<10 then begin
 if length(fecha) = 7 then begin
   dia:=('0' + fecha[1]);
   mes:=(fecha[3] + fecha[4]);
   ano:=('20' + fecha[6] + fecha[7]);
   fechalarga:=dia + '/' + mes + '/' + ano;
 end
 else begin
   dia:=(fecha[1] + fecha[2]);
   mes:=(fecha[4] + fecha[5]);
   ano:=('20' + fecha[7] + fecha[8]);
   fechalarga:=dia + '/' + mes + '/' + ano;
 end;
end
else begin
 fechalarga:=fecha;
end;

end;

function separar(cadena:string;separador:string;posi:integer):string;
var cad1:string;
    x,y,kl1,kl2:integer;
    fd:boolean;
begin
 fd:=true;x:=1;y:=0;
 while fd do begin
    if cadena[x]=separador then begin
       y:=y+1;

       if y = 1 then kl1:=x;
       {Opcional para encontrar la primera coma solamente
       if y = 1 then begin
         kl1:=x;fd:=false;
       end;  }
       if y = 2 then begin
         kl2:=x;fd:=false;
       end;
    end;
    if cadena[x] = chr(39) then cadena[x]:='"';
    x:=x+1;
 end;
 cad1:='';
 case posi of
   1: for y:=1 to kl1-1 do
     cad1:=cad1 + cadena[y];
   2: for y:=kl1+1 to kl2-1 do
     cad1:=cad1 + cadena[y];
   3: for y:=kl2+1 to length(cadena) do
     cad1:=cad1 + cadena[y];
 end;
 separar:=cad1;
end;

Function StrFill(cadena:string;longitud:Integer;relleno,posicion:string): string;
var cTemp:string;
    n, nVeces:Integer;
begin
cTemp := cadena;
nVeces := longitud - length(cTemp);
for n := 1 to nVeces do begin
  if ((posicion = 'D') or (posicion = 'd')) then
    cTemp := cTemp+relleno
  else
    cTemp := relleno+cTemp;
end;
Result := cTemp;
end;

Function cortastr(cadena:string;tamano:integer):lineas;
var y:integer;
    cad1,cad2:string;
    bandera:boolean;
begin
 if length(cadena) < (tamano+1) then begin
   cortastr[1]:=cadena;cortastr[2]:='';
 end
 else begin
   bandera:=false;
   for y:=1 to length(cadena) do begin
     if (y > (tamano - 5)) and (cadena[y] = ' ') then bandera:=true;
     if bandera then cad2:=cad2+cadena[y]
     else cad1:=cad1+cadena[y];
   end;
   cortastr[1]:=cad1;cortastr[2]:=cad2;
 end;
end;

Function strcut(cadena:string;tamano:integer):string;
var x:integer;
    hj:string;
begin
hj:='';
if length(cadena)< tamano then
  for x:=1 to length(cadena) do hj:=hj+cadena[x]
else for x:=1 to tamano do hj:=hj+cadena[x];
strcut:=hj;
end;

Function cuenta(cadena:widestring;separador:string):integer;
var xl,jk:integer;
begin
jk:=0;
for xl:=1 to length(cadena) do
   if cadena[xl]=separador then jk:=jk+1;
cuenta:=jk;
end;



function tomadato(cadena:string;lugar:integer):integer;
var cad1:string;
    x,pod,y,ant,tot:integer;
    fd:boolean;
begin
 y:=0;fd:=true;x:=1;ant:=1;pod:=1;
 while fd do begin
    if cadena[x]=',' then begin
      y:=y+1;
      if lugar=y then begin
        pod:=x-1;fd:=false;
      end
      else begin
        if lugar=1 then ant:=1
        else ant:=x+1;
      end;
    end
    else begin
      if (lugar = 7) and (y = 6) then begin
         ant:=x;pod:=length(cadena);fd:=false;
      end;
    end;
    x:=x+1;
    if x > length(cadena) then fd:=false;
 end;
 tot:=pod-ant+1;
 for x:=1 to tot do begin
   if x=1 then cad1:=cadena[ant]
   else cad1:=cad1 + cadena[(ant + x - 1)];
 end;
 tomadato:=strtoint(trim(cad1));
end;


function Extrae(cadena:widestring;separador:char;pos:integer):string;
var cad1:string;
    maxseparador,posini,posfin,uy:integer;
begin
 maxseparador:=cuenta(cadena,separador);
 posini:=BuscaChr(cadena,separador,pos-1);
 if pos > maxseparador then posfin:=length(cadena)
 else posfin:=BuscaChr(cadena,separador,pos)-2;
 Result:=strcut_midle(cadena,posini,posfin);
end;

{Funcion que busca el caracter y devuelve la posision que lo encuentra}
function BuscaChr(cadena:widestring;caracter:string;posi:integer):integer;
var ht,cb:integer;
begin
if posi = 0 then Result:=1
else begin
  cb:=0;
  for ht:=1 to length(cadena) do begin
    if cadena[ht] = caracter then cb:=cb + 1;
    if cb = posi then begin
      Result:=ht+1;
      Break;
    end;
  end;
end;
end;

function separar1(cadena:string;separador:string):string;
var cad1:string;
    x,y,tot:integer;
    fd:boolean;
begin
 fd:=true;x:=1;
 while fd do begin
    if cadena[x]=separador then fd:=false;
    x:=x+1;
 end;
 tot:=length(cadena)-x+1;
 for y:=1 to tot do begin
   if y=1 then cad1:= cadena[x]
   else cad1:=cad1 + cadena[(x + y - 1)];
 end;
 separar1:=cad1;
end;

{El tamaño debe incluir la coma}
function incorpora(cadena:string;tamano,decimal:integer):string;
var x,y,pos:integer;
    cad1,parte1,parte2:string;
begin
setlength(cad1,tamano);{Definimos el tamaño del string}
for x:=1 to tamano do cad1[x]:='0';{Llenamos el string de '0'}
if (strtofloat(trim(cadena))-trunc(strtofloat(trim(cadena))))> 0 then begin
{El numero es decimal}
  for x:=1 to length(trim(cadena)) do
     if cadena[x] = ',' then pos:=x;
  for x:=1 to pos-1 do parte1:=parte1+cadena[x];{Parte entera}
  for x:=pos+1 to length(trim(cadena)) do parte2:=parte2+cadena[x];{Parte decima}
  cad1[tamano-decimal]:=',';{Colocamos la coma}
  y:=length(parte1);
  for x:=(tamano-decimal-1) downto 1 do begin
    {de la coma para la izquierda}
    if y > 0 then cad1[x]:=parte1[y];
    y:=y-1;
  end;
  y:=1;
  for x:=(tamano-decimal+1) to tamano do begin
    {de la coma para la derecha}
    if y < (length(parte2)+1) then cad1[x]:=parte2[y];
    y:=y+1;
  end;
  incorpora:=cad1;
end
else begin
{El numero es entero}
  parte1:=cadena;{Parte entera}
  cad1[tamano-decimal]:=',';{Colocamos la coma}
  y:=length(parte1);
  for x:=(tamano-decimal-1) downto 1 do begin
    {de la coma para la izquierda}
    if y > 0 then cad1[x]:=parte1[y];
    y:=y-1;
  end;
  incorpora:=cad1;
end;
end;

{1=año;2=mes;3=dia}
function mda(fecha:string;tipo:byte):integer;
var lafecha:string ;
begin
lafecha:=fechalarga(fecha);
if tipo=1 then begin{devuelve el año}
   Result:=strtoint(lafecha[7]+lafecha[8]+lafecha[9]+lafecha[10]);
end;
if tipo=2 then begin{devuelve el mes}
  Result:=strtoint(lafecha[4]+lafecha[5]);
end;
if tipo=3 then begin{devuelve el dia}
  Result:=strtoint(lafecha[1]+lafecha[2]);
end;
end;

function crea_cuim(nro_dni,prefijo:integer):string;
var cuim,eldni,constante,postfijo:string;
    m1,m2:string;
    tc,nSuma:integer;
begin
 eldni:=llenastr(inttostr(nro_dni),8,'0');
 nSuma:=0;
 constante:='5432765432';
 cuim:=llenastr(inttostr(prefijo),2,'0') + '-' + eldni;
 for tc := 10 downto 1 do begin
   if tc < 3 then begin
     m1:=cuim[tc];m2:=constante[tc];
     nSuma:=nSuma + (strtoint(m1)*strtoint(m2))
   end
   else begin
     m1:=cuim[tc+1];m2:=constante[tc];
     nSuma:=nSuma + (strtoint(m1)*strtoint(m2))
   end;
 end;
 if (nSuma mod 11) = 0 then postfijo:='0'
 else begin
   if (nSuma mod 11) > 1 then
        postfijo:=inttostr(11 - (nSuma mod 11))
   else postfijo:='1';
 end;
 crea_cuim:=cuim + '-' + postfijo;
end;

{Funcion que completa los numeros con un determinado caracter para formar
00023456}
Function llenastr(cadena:string;tamano:integer;caracter:char):string;
var x,fd:integer;
    cad1:string;
begin
if length(cadena)=tamano then llenastr:=cadena
else begin
   setlength(cad1,tamano);{Definimos el tamaño del string}
   //y:=0;
   fd:=length(cadena);
   for x:=tamano downto 1 do begin
      {if y = 3 then begin
         cad1[x]:='.';
         y:=-1;
      end
      else begin}
        if fd < 1 then cad1[x]:=caracter
        else cad1[x]:=cadena[fd];
        fd:=fd-1;
      {end;}
      //y:=y+1;
   end;
   llenastr:=cad1;
end;
end;

Function Cambia(cadena:string;caracter1,caracter2:char):string;
var sd:integer;
begin
for sd:=1 to length(cadena) do
  if cadena[sd]=caracter1 then cadena[sd]:=caracter2;
cambia:=cadena;
end;

function inttocuim(nro_cuim:integer):string;
var cadena,cad1:string;
    sd,sf:integer;
begin
if length(inttostr(nro_cuim)) <> 11 then inttocuim:='00-00000000-0'
else begin
  setlength(cadena,13);
  cad1:=inttostr(nro_cuim);
  sf:=1;
  for sd:=1 to 13 do begin
     if (sd = 3) or (sd = 12) then cadena[sd]:='-'
     else begin
       cadena[sd]:=cad1[sf];
       sf:=sf+1;
     end;
  end;
  inttocuim:=cadena;
end;
end;

function cuimtoint(nro_cuim:string):string;
var cadena:string;
    sd,sf:integer;
begin
if length(nro_cuim) <> 13 then cuimtoint:='0'
else begin
  setlength(cadena,11);
  sf:=1;
  for sd:=1 to 13 do begin
     if nro_cuim[sd]<> '-' then begin
        cadena[sf]:=nro_cuim[sd];
        sf:=sf+1;
     end;
  end;
  cuimtoint:=cadena;
end;
end;


function inttodni(nro_cuim:string):string;
var cadena:string;
    sd,sf:integer;
begin
if length(nro_cuim)<> 13 then inttodni:='00000000'
else begin
   setlength(cadena,8);
   sf:=1;
   for sd:=4 to 11 do begin
      cadena[sf]:=nro_cuim[sd];
      sf:=sf+1;
   end;
   inttodni:=cadena;
end;
end;


Procedure AuditAccesos(AreaIngreso:string;OperacionRealizada:char);
var mbn:string;
begin
 try
  if StrToInt(modulo2.UsuNro) <> 0 then begin
   if trim(modulo2.num_modulo) = '' then mbn:='NULL'
   else mbn:=modulo2.num_modulo;
   ComUni:=TADOCommand.Create(nil);
   ComUni.ConnectionString:=modulo2.ConectDbMuni;
   ComUni.CommandText:='INSERT INTO Accesos (codusu,area,operacion,idmodulo,'
   + 'pv,Host,IP) VALUES (' + modulo2.UsuNro + ' , ' + chr(39) + strcut(AreaIngreso,15) + chr(39)
   + ' , ' + chr(39) + OperacionRealizada + chr(39) + ' , ' + mbn + ' , ' + modulo2.UsuPuntoV
   + ' , ' + chr(39) + modulo2.MyHost + chr(39) + ' , ' + chr(39) + modulo2.MyIP + chr(39) + ');';
   ComUni.Execute;
   ComUni.Free;
  end;
 except
 end;
end;

Function AuditaImpresos(OperacionRealizada,ElCuim:string):string;
begin
ComUni:=TADOCommand.Create(nil);
ComUni.ConnectionString:=modulo2.ConectDbMuni;
DtsUni:=TADODataSet.Create(nil);
DtsUni.ConnectionString:=modulo2.ConectDbSeguridad;
ComUni.CommandText:='INSERT INTO Impresos (codusu,Listado,idmodulo,pv,CUIM) VALUES ('
+ modulo2.UsuNro + ' , ' + chr(39) + strcut(OperacionRealizada,40) + chr(39)
+ ' , ' + modulo2.num_modulo + ' , ' + modulo2.UsuPuntoV + ' , ' + chr(39) + trim(ElCuim) + chr(39) + ');';
ComUni.Execute;
ComUni.CommandText:='SELECT @@IDENTITY AS Id';
DtsUni.Recordset:=ComUni.Execute;
DtsUni.Open;
Result:=DtsUni.Fields[0].AsString;
DtsUni.Close;
DtsUni.Free;
end;

Function fechacompleta(fecha:string):string;
var dia,mes,ano:string;
begin
if fecha ='' then fechacompleta:=''
else begin
  if (length(fecha)<10) then begin
   if length(fecha) = 7 then begin
     dia:=('0' + fecha[1]);
     mes:=(fecha[3] + fecha[4]);
     ano:=('19' + fecha[6] + fecha[7]);
     fechacompleta:=dia + '/' + mes + '/' + ano;
   end
   else begin
     dia:=(fecha[1] + fecha[2]);
     mes:=(fecha[4] + fecha[5]);
     ano:=('19' + fecha[7] + fecha[8]);
     fechacompleta:=dia + '/' + mes + '/' + ano;
   end;
  end
  else begin
   fechacompleta:=fecha;
  end;
end;
end;


function strtocuim(nro_cuim:string):string;
var cadena,cad1:string;
    sd,sf:integer;
begin
if length(nro_cuim) <> 11 then strtocuim:='00-00000000-0'
else begin
  setlength(cadena,13);
  cad1:=nro_cuim;
  sf:=1;
  for sd:=1 to 13 do begin
     if (sd = 3) or (sd = 12) then cadena[sd]:='-'
     else begin
       cadena[sd]:=cad1[sf];
       sf:=sf+1;
     end;
  end;
  strtocuim:=cadena;
end;
end;

Function Fec_System():string;
begin
DtsUni:=TADODataSet.Create(nil);
DtsUni.ConnectionString:=modulo2.ConectDbMuni;
DtsUni.CommandText:='FECHA';
DtsUni.Open;
Fec_System:=Corta_fec(DtsUni.Fields[0].AsString);
DtsUni.Close;
DtsUni.Free;
end;

Function Corta_fec(cadena:string):string;
var sd:integer;
    cad1:string;
begin
if length(cadena) < 16 then Corta_fec:=cadena
else begin
for sd:=1 to 16 do cad1:= cad1 + cadena[sd];
Corta_fec:=cad1;
end;
end;


Function V_Modulos(cadena:string):boolean;
begin
DtsUni:=TADODataSet.Create(nil);
DtsUni.ConnectionString:=modulo2.ConectDbMuni;
DtsUni.CommandText:='SELECT Nivel FROM Usumodulos WHERE (((codusu)='
     + modulo2.UsuNro + ')) AND (((modulo) = ' + cadena + '));';
try
  DtsUni.Open;
  if DtsUni.RecordCount <> 0 then begin
    V_Modulos:=true;
    modulo2.num_modulo:=cadena;
    modulo2.UsuNvl:=DtsUni.Fields[0].AsInteger;
  end
  else V_Modulos:=false;
  DtsUni.Close;
except  begin
  V_Modulos:=true;
  modulo2.UsuNvl:=3;
  DtsUni.Close;
  end;
end;
DtsUni.Free;
end;

function calcular_deuda(fec_ant,fec_act:string):real;
var int_mes,a,m,d,totmes,dias_mes:integer;
var int_dia,total:real;
begin
 int_mes:=1;int_dia:=0.03287671232876712328;
 case mda(fec_act,2) of
   1,3,5,7,8,10,12: dias_mes:=31;
                 2: dias_mes:=28;
          4,6,9,11: dias_mes:=30;
 end;
 a:=mda(fec_act,1)-mda(fec_ant,1);
 m:=mda(fec_act,2)-mda(fec_ant,2);
 d:=mda(fec_act,3)-mda(fec_ant,3);
 totmes:=(a*12)+m;
 if d < 0 then begin
   totmes:=totmes-1;
   d:=dias_mes+d;
 end;
 total:=(totmes*int_mes)+(d*int_dia);
 if total < 0 then calcular_deuda:=0
 else calcular_deuda:=total;
end;

function decimales(cadena:string;lugares:byte):real;
var total:real;
    resul:string;
    sd:integer;
    lk:LongInt;

begin
lk:=1;
for sd:=1 to lugares do lk:=lk*10;
total:=strtofloat(cadena);
total:=total*lk;
resul:=floattostr(total);
for sd:=1 to length(resul) do begin
   if (resul[sd]=chr(44)) or (resul[sd]=chr(46)) then begin
     resul:=strcut(resul,sd-1);
   end;
end;
total:=strtofloat(resul)/lk;
decimales:=total;
end;

function redondeo(cadena:string;decimales:integer):string;
var el_numero:real;
begin
try
  el_numero:=strtofloat(cambia(cadena,chr(44),chr(46)));
except
  el_numero:=strtofloat(cambia(cadena,chr(46),chr(44)));
end;
redondeo:=FloatToStrF(el_numero, ffFixed, 20, decimales);
end;



Function strcut_midle(cadena:widestring;comienzo,fin:integer):string;
var x:integer;
    hj:string;
begin
hj:='';
if fin > length(cadena) then fin := length(cadena);
if fin = 0 then fin := length(cadena);
for x:=comienzo to fin do hj:=hj+cadena[x];
strcut_midle:=hj;
end;


Function completa1(cadena:string;tamano:integer;caracter:char):string;
var x,fd:integer;
    cad1:string;
begin
 if length(cadena)=tamano then completa1:=cadena
 else begin
  setlength(cad1,tamano);{Definimos el tamaño del string}
  fd:=length(cadena);
  for x:=tamano downto 1 do begin
   if fd < 1 then cad1[x]:=caracter
    else cad1[x]:=cadena[fd];
   fd:=fd-1;
  end;
  completa1:=cad1;
 end;
end;

{Funcion que completa un string con caracteres de acuerdo al formato del banco,
tambien elimina la coma. Decimales debe ser mayor a 0(cero)}
function importetobanco(cadena:string;tamano,decimales:integer;caracter:char):string;
var x,y,fd:integer;
    cad1,cad2,cad3:string;
begin

  cad3:=redondeo(cadena,decimales);{Redondeamos para tener el valor con los decimales designados}

  setlength(cad1,tamano-decimales);{Definimos el tamaño del string}
  setlength(cad2,decimales);{Definimos el tamaño del string}
  y:=0;
  if (findchr(cad3,',') = 0) and (findchr(cad3,'.') = 0) then cad3:=cad3 + '.00';

  for x:=1 to length(cad3) do begin
   if (cad3[x]='.') or (cad3[x]=',') then y:=x;
  end;
  fd:=y-1;
  for x:=(tamano-decimales) downto 1 do begin
   if fd < 1 then cad1[x]:=caracter
   else cad1[x]:=cad3[fd];
   fd:=fd-1;
  end;
  fd:=1;
  for x:=y+1 to (y+decimales) do begin
   if x > length(cad3) then cad2[fd]:=caracter
   else cad2[fd]:=cad3[x];
   fd:=fd+1
  end;
  importetobanco:=cad1+cad2;
end;

function fecha_dif(f_desde,f_hasta:string):string;
var an,an1,ms,ms1,ds,ds1:integer;
var fromdate,todate:TDateTime;
begin
an:=mda(f_desde,1);
ms:=mda(f_desde,2);
ds:=mda(f_desde,3);
an1:=mda(f_hasta,1);
ms1:=mda(f_hasta,2);
ds1:=mda(f_hasta,3);
fromdate:=EncodeDateTime(an,ms,ds,0,0,0,0);
todate:=EncodeDateTime(an1,ms1,ds1,0,0,0,0);
fecha_dif:=floattostr(DaySpan(todate,fromdate)+1);
end;

function digito(cadena:string;opc:byte):string;
var t,q,total,gr:integer;
    cad,cad2:string;
begin
  q:=Length(cadena);
  cad2:='3579';
  cad:='1';
  gr:=1;
  for t:=1 to q do begin
    cad:=cad + cad2[gr];
    if gr = 4 then gr:=0;
    gr:=gr + 1;
  end;
  cad2:=cadena;total:=0;
  for t:=1 to q do
     total:= total + (strtoint(cad2[t])*strtoint(cad[t]));
  gr:=trunc(total/2);
  gr:=gr mod 10;
  digito:=inttostr(gr);
end;

function codigo_barra(cadena:string):string;
var v_star,v_stop,lcheck:char;
    v_check,v_asc,v_largo,ro,v_num:integer;
    la_cadena,la_ret:string;
begin
  v_star := chr(105 + 32);
  v_stop := chr(106 + 32);
  v_check := ord(v_star)-32;
  la_cadena := trim(cadena);
  v_largo := length(la_cadena);
  {La longitud debe ser par}
  if (v_largo mod 2) <> 0 then begin
     la_cadena:= '0' + la_cadena;
     v_largo:= length(la_cadena);
  end;
  {Convierto los pares a caracteres}
  la_ret:='';
  ro:=2;
  while ro < (v_largo + 2) do begin
    v_num:=strtoint(la_cadena[ro-1] + la_cadena[ro]) + 32;
    la_ret:=la_ret + chr(v_num);
    ro:=ro+2;
  end;
  v_largo:=length(la_ret);
  for ro:=1 to v_largo do begin
    v_num:=ord(la_ret[ro])-32;
    v_asc:=(v_num * ro);
    v_check:=v_check + v_asc;
  end;
  ro:=(v_check mod 103) + 32;
  lcheck:= chr(ro);
  la_ret:=v_star + la_ret + lcheck + v_stop;
  {Esta parte cambia los espacios y caracteres invalidos}
  la_ret:=cambia(la_ret,chr(32),chr(232));
  la_ret:=cambia(la_ret,chr(127),chr(192));
  la_ret:=cambia(la_ret,chr(128),chr(193));
  codigo_barra:=la_ret;
end;

{Busca el caracter en la cadena y devuelve la cantidad de veces que lo encuentra}
function findchr(cadena:string;caracter:char):integer;
var xf,vec:integer;
begin
vec:=0;
for xf:=1 to length(cadena) do
  if cadena[xf] = caracter then vec:=vec+1;
findchr:=vec;
end;


Function Fec_Texto(cadena:string):string;
var la_fecha,dnc,mnc,ano:string;
begin
  la_fecha:=(cadena);
  if length(la_fecha) <> 10 then la_fecha:=fechalarga(cadena);
  dnc:=la_fecha[1]+la_fecha[2];
  mnc:=la_fecha[4]+la_fecha[5];
  ano:=la_fecha[7]+la_fecha[8]+la_fecha[9]+la_fecha[10];
  mnc:=el_mes(strtoint(mnc));
  Fec_texto:=dnc + ' de ' + mnc + ' del ' + ano;
end;

function el_mes(mes:integer):string;
begin
  if (mes > 12) or (mes < 1) then el_mes:='Desconocido'
  else begin
    if mes = 1 then el_mes:='Enero';
    if mes = 2 then el_mes:='Febrero';
    if mes = 3 then el_mes:='Marzo';
    if mes = 4 then el_mes:='Abril';
    if mes = 5 then el_mes:='Mayo';
    if mes = 6 then el_mes:='Junio';
    if mes = 7 then el_mes:='Julio';
    if mes = 8 then el_mes:='Agosto';
    if mes = 9 then el_mes:='Setiembre';
    if mes = 10 then el_mes:='Octubre';
    if mes = 11 then el_mes:='Noviembre';
    if mes = 12 then el_mes:='Diciembre';
  end;
end;

{Fncion que determina el vencimiento de la ultima cuota}
Function fechaCuotas(cadena:string;cuotas:integer):string;
var valfecha:string;
begin
valfecha:=DateTimetoStr(IncMonth(strtodatetime(cadena),cuotas));
fechaCuotas:='10' + strcut_midle(valfecha,3,10);
end;


{Funcion que toma una fecha y elimina los '/' para dejarla como numero}
function fectoint(cadena:string):string;
var lw:integer;
    cad1:string;
begin
cad1:='';
for lw:=1 to length(cadena) do begin
  if lw <> 3 then begin
    if (lw <> 6) then cad1:=cad1 + cadena[lw];
  end;
end;
fectoint:=cad1;
end;

{Valor x elevado a la n}
Function Potencia(x,n:Real):Real;
begin
Potencia := Exp(n *Ln(x))
end;

{el formato del numero ya tiene que tener 2 decimales}
function numtopesos(num:string):string;
  var centena_,decena_,unidad_,diez_,veinte_: StrArray;
  var grupo,entero,decimales,letra:string;
  var cen_,dec_,uni_,contador,incremento:integer;
begin
  {Declaramos los valores de los arrays}
  centena_[1]:='CIENTO';centena_[2]:='DOSCIENTOS';centena_[3]:='TRESCIENTOS';
  centena_[4]:='CUATROCIENTOS';centena_[5]:='QUINIENTOS';centena_[6]:='SEISCIENTOS';
  centena_[7]:='SETECIENTOS';centena_[8]:='OCHOCIENTOS';centena_[9]:='NOVECIENTOS';

  decena_[1]:='DIEZ';decena_[2]:='VEINTE';decena_[3]:='TREINTA';
  decena_[4]:='CUARENTA';decena_[5]:='CINCUENTA';decena_[6]:='SESENTA';
  decena_[7]:='SETENTA';decena_[8]:='OCHENTA';decena_[9]:='NOVENTA';

  unidad_[1]:='UNO';unidad_[2]:='DOS';unidad_[3]:='TRES';
  unidad_[4]:='CUATRO';unidad_[5]:='CINCO';unidad_[6]:='SEIS';
  unidad_[7]:='SIETE';unidad_[8]:='OCHO';unidad_[9]:='NUEVE';

  diez_[1]:='ONCE';diez_[2]:='DOCE';diez_[3]:='TRECE';
  diez_[4]:='CATORCE';diez_[5]:='QUINCE';diez_[6]:='DIECISEIS';
  diez_[7]:='DIECISIETE';diez_[8]:='DIECIOCHO';diez_[9]:='DIECINUEVE';

  veinte_[1]:='VEINTIUNO';veinte_[2]:='VEINTIDOS';veinte_[3]:='VEINTITRES';
  veinte_[4]:='VEINTICUATRO';veinte_[5]:='VEINTICINCO';veinte_[6]:='VEINTISEIS';
  veinte_[7]:='VEINTISIETE';veinte_[8]:='VEINTIOCHO';veinte_[9]:='VEINTINUEVE';

  {Preparo el numero}
  {El número debe tener decimales por lo tanto si no los tiene,
  se los agrego}
  if (findchr(num,chr(44)) = 0) and (findchr(num,chr(46)) = 0) then num:=num + ',00';
  entero:=completa1(num,15,'0');
  entero:=completa2(num,12,2,'0');
  decimales:=strcut_midle(entero,14,15);
  entero:=strcut_midle(entero,1,12);
  letra:='';
  if strtoint(entero)=0 then begin
    numtopesos:='CERO.---';
    exit;
  end;

  {Tenemos que tomar de cifras de a tres la parte entera}
  contador:=1;incremento:=3;
  while (contador < length(entero)) do begin
    grupo:=strcut_midle(entero,contador,contador+incremento-1);
    cen_:=strtoint(strcut_midle(grupo,1,1));
    dec_:=strtoint(strcut_midle(grupo,2,2));
    uni_:=strtoint(strcut_midle(grupo,3,3));
    {Centena}
    if cen_ <> 0 then begin
      {Exepcion de cien con ciento}
      if strtoint(grupo) = 100 then letra:=letra + ' CIEN'
      else letra:=letra + ' ' + centena_[cen_];
    end;
    {Las decenas y las exepciones}
    case dec_ of
      1: begin   {Decenas de Diez}
           if uni_ = 0 then letra:=letra + ' ' + decena_[dec_]
           else letra:=letra + ' ' + diez_[uni_];
         end;
      2: begin  {Decenas de Veinte}
           if uni_ = 0 then letra:=letra + ' ' + decena_[dec_]
           else letra:=letra + ' ' + veinte_[uni_];
         end;
    else begin {Otras decenas}
      if dec_ <> 0 then letra:=letra + ' ' + decena_[dec_];
      {Unidad}
      if uni_ <> 0 then begin
        if dec_ <> 0 then letra:=letra + ' Y ' + unidad_[uni_]
        else letra:=letra + ' ' + unidad_[uni_];
      end;
     end;
    end;{end del Case}
    {Caso del Mil, Millon con el uno}
    if (uni_ = 1) and ((contador = 1) or (contador = 4) or (contador = 7)) and (dec_ <> 1) then
      letra:=strcut_midle(letra,1,length(letra)-1);
    {Puntos del mil y millon}
    if (contador = 4) and (strtoint(grupo) > 0) then begin
      if (strtoint(grupo)=1) and (strtoint(strcut_midle(entero,1,3)) = 0) then letra:=letra + ' MILLON'
      else letra:=letra + ' MILLONES';
    end;
    if ((contador = 1) or (contador = 7)) and (strtoint(grupo) > 0) then begin
      letra:=letra + ' MIL';
      {Exepciones del mil de millon si millon es cero}
      if (contador = 1) and (strtoint(strcut_midle(entero,4,6)) = 0) then letra:=letra + ' MILLONES';
    end;
    contador:=contador+incremento;
  end;
  {Parte de los decimales}
  if strtoint(decimales)>0 then begin
    dec_:=strtoint(strcut_midle(decimales,1,1));
    uni_:=strtoint(strcut_midle(decimales,2,2));
    case dec_ of
      1: begin
           if uni_ = 0 then letra:=letra + ' CON ' + decena_[dec_]
           else letra:=letra + ' CON ' + diez_[uni_];
         end;
      2: begin
           if uni_ = 0 then letra:=letra + ' CON ' + decena_[dec_]
           else letra:=letra + ' CON ' + veinte_[uni_];
         end;
    else begin {Otras decenas}
      if dec_ <> 0 then letra:=letra + ' CON ' + decena_[dec_];
      {letra:=letra + ' CON';}
      if uni_ <> 0 then begin
        if dec_ = 0 then letra:=letra + ' CON ' + unidad_[uni_]
        else letra:=letra + ' Y ' + unidad_[uni_];
      end;
      end;
    end;{End del case}
    if (uni_ = 1) and (dec_ <> 1) then letra:=strcut_midle(letra,1,length(letra)-1);
    letra:=letra + ' Cvos.';
  end;
  numtopesos:=letra;
end;

{Funcion que completa los numeros con un determinado caracter para formar
00023456.90}
Function completa2(cadena:string;tamano1,tamano2:integer;caracter:char):string;
var x,y,fd:integer;
    cad1,cad2:string;
    cad3:char;
begin
setlength(cad1,tamano1);{Definimos el tamaño del string}
setlength(cad2,tamano2+1);{Definimos el tamaño del string}
y:=0;
for x:=1 to length(cadena) do begin
  if (cadena[x]='.') or (cadena[x]=',') then y:=x;
  if (cadena[x]='.') then cad3:='.';
  if (cadena[x]=',') then cad3:=',';
end;
if y<>0 then begin
  fd:=y-1;
  for x:=tamano1 downto 1 do begin
    if fd < 1 then cad1[x]:=caracter
    else cad1[x]:=cadena[fd];
    fd:=fd-1;
  end;
  cad2[1]:=cad3;fd:=2;
  for x:=y+1 to (y+tamano2) do begin
    if x > length(cadena) then cad2[fd]:=caracter
    else cad2[fd]:=cadena[x];
    fd:=fd+1
  end;
end
else begin
  fd:=length(cadena);
  for x:=tamano1 downto 1 do begin
    if fd < 1 then cad1[x]:=caracter
    else cad1[x]:=cadena[fd];
    fd:=fd-1;
  end;
  for x:=1 to tamano2+1 do begin
    if x=1 then cad2[x]:='.'
    else cad2[x]:=caracter;
  end;
end;
completa2:=cad1+cad2;
end;

{Cambia el caracter 'carac1' por el caracter 'carac2'}
function transforma(cadena:string;carac1,carac2:char):string;
var lw:integer;
    ou:string;
    ft:real;
begin
try
  ft:=strtofloat(cadena);
  ou:=floattostr(ft);
except
  ou:=cadena;
end;
for lw:=1 to length(ou) do
   if ou[lw]= carac1 then ou[lw]:=carac2;
transforma:=ou;
end;

{Fncion que Incremeta la Fecha en Meses}
Function MesIncremento(Fecha:string;Meses:integer):string;
begin
Result:=DateTimetoStr(IncMonth(strtodatetime(Fecha),Meses));
end;

{Fncion que Incremeta la Fecha en Dias}
Function DiaIncremento(Fecha:string;Dias:integer):string;
begin
Result:=DateTimetoStr(IncDay(strtodatetime(Fecha),Dias));
end;

{Fncion que Incremeta la Fecha en Dias}
Function AnioIncremento(Fecha:string;Anios:integer):string;
begin
Result:=DateTimetoStr(IncYear(strtodatetime(Fecha),Anios));
end;

function DiasDiferencia(FechaInico,FechaFin:string):integer;
var FecIni,FecFin:TDateTime;
begin
FecIni:=StrToDateTime(FechaInico);
FecFin:=StrToDateTime(FechaFin);
Result:=DaysBetween(FecIni,FecFin);
end;

Procedure EnviarLog(NombreVentana:string;Texto:widestring);
var ru1:string;
       fichero1:textfile;
       ty:char;rt,hf,hg:byte;
begin
ru1:=modulo2.path_todo + '\Datos\Log.txt';
assignfile(fichero1,ru1);
if not FileExists(ru1) then Rewrite(fichero1)
else Append(fichero1);
Writeln(fichero1,NombreVentana + ' - ' + DateTimeToStr(now()) + ' - ' + Texto);
closefile(fichero1);
end;


{Funcion que de acuerdo al valor de tpo, encripta o desencripta la cadena
recibida.
El proceso de encriptacion cuenta con la mezcla de caracteres, suma de un valor
especifico y la integracion de cuatro caracteres numericos al azar en
determinada posicion. Las posiciones de integracion son 4,8,16,24.
La mezcla es la intercalacion del primero con el segundo, el tercero con el
cuarto... El valor a sumar es 1}
function encripta(cadena:string;tpo:boolean):string;
 var lw,kw:integer;
 var resul1,resul2,c1,c2:string;
begin
 if tpo then begin
  {Valor TRUE es encriptacion de la cadena}
  {Empezamos con la suma del valor 1 a todos los caracteres}
  resul1:='';{esta variable tiene el resultado}
  for lw:=1 to length(cadena) do begin
    if cadena[lw]='9' then begin
      resul1:= resul1 + '0'; {caso que la suma da 10}
    end
    else begin
      kw:=strtoint(cadena[lw])+1;{le sumamos 1}
      resul1:=resul1 + inttostr(kw);
    end;
  end;
  {Empezamos con la mezcla de datos}
  lw:=2;resul2:='';
  while not(lw > length(cadena)) do begin
     c1:=resul1[lw-1];{mezclamos el caracter 1 con el 2}
     c2:=resul1[lw];
     resul2:=resul2 + c2 + c1;
     lw:=lw+2;
  end;
{Empezamos con generar numeros al azar y agregarlos en las posiciones indicadas}
  resul1:='';kw:=0;
  setlength(resul1,length(cadena)+4);{Definimos el tamaño final del string}
  c1:=inttostr(random(9)) + inttostr(random(9)) + inttostr(random(9)) + inttostr(random(9));
  for lw:=1 to length(resul2) do begin
     if (lw = 4) or (lw = 8) or (lw = 16) or (lw = 24) then begin
        resul1[lw+kw]:=c1[kw+1];
        resul1[lw+kw+1]:=resul2[lw];
        kw:=kw+1;
     end
     else resul1[lw+kw]:=resul2[lw];
  end;
  encripta:=resul1;
 end
 else begin
{Empezamos a desencriptar la cadena ingresada}
{Empezamos con la eliminacion de los valores al azar incorporados}
  resul1:='';kw:=1;
  setlength(resul1,length(cadena)-4);{Definimos el tamaño original del string}
  for lw:=1 to length(cadena) do begin
     if (lw = 4) or (lw = 9) or (lw = 18) or (lw = 27) then
     else begin
       resul1[kw]:=cadena[lw];
       kw:=kw+1;
     end;
  end;
  {Empezamos a desmezclar los numeros}
  lw:=2;resul2:='';
  while not(lw > length(resul1)) do begin
     c1:=resul1[lw-1];{mezclamos el caracter 1 con el 2}
     c2:=resul1[lw];
     resul2:=resul2 + c2 + c1;
     lw:=lw+2;
  end;
  {Empezamos a restar el valor sumado}
  resul1:='';{esta variable tiene el resultado}
  for lw:=1 to length(resul2) do begin
    if resul2[lw]='0' then begin
      resul1:= resul1 + '9'; {caso que la resta de -1}
    end
    else begin
      kw:=strtoint(resul2[lw])-1;{le restamos 1}
      resul1:=resul1 + inttostr(kw);
    end;
  end;
  encripta:=resul1;
 end;
end;

procedure EncontrarHijos(Codigo,CantTabs,PosRec:integer;Tabla:TADOTable);
var Cod,CntTab,gr,PosRecNo:integer;
begin
for gr:=PosRec to Tabla.RecordCount do begin
  Tabla.RecNo:=gr;
  if (Tabla.Fields[1].AsInteger = Codigo) then begin
    PosRecNo:=Tabla.RecNo;
    Cod:=Tabla.Fields[0].AsInteger;
    CntTab:=CantTabs+1;
    ListaOrganigrama.Add(StrFill(inttostr(CntTab),4,'0','') + trim(Tabla.Fields[2].AsString)
    +' - (' + StrFill(Tabla.Fields[0].AsString,4,'0','')+')');
    EncontrarHijos(Cod,CntTab,PosRecNo,Tabla);
  end;
end;
end;


Procedure CrearTreeConTbs(FileName,NombreTabla:string;Tabla:TADOTable);
var txt:string;
   Texto:widestring;
   fichero1:textfile;
   gj,k:integer;
begin
if FileExists(FileName) then DeleteFile(FileName);
assignfile(fichero1,FileName);
Rewrite(fichero1);
if Tabla.RecordCount > 0 then begin
  for gj:=1 to Tabla.RecordCount do begin
    Tabla.RecNo:=gj;
    txt:='';
    for k:=1 to Tabla.Fields[1].AsInteger do begin
      txt:=txt + chr(9);
    end;
    Texto:=txt + trim(Tabla.Fields[2].AsString) ; // + ' - (' + StrFill(Tabla.Fields[0].AsString,4,'0','') + ')'
    Writeln(fichero1,Texto);
  end;
end
else Writeln(fichero1,'SIN ORGANIGRAMA');
try
  closefile(fichero1);
//  closefile(fichero1);
except
end;
end;

function usu_level(usu,area:integer):integer;
begin
try
  DtsUni:=TADODataSet.Create(nil);
  DtsUni.ConnectionString:=modulo2.ConectDbSeguridad;
  DtsUni.CommandText:='SELECT Nivel FROM UsuModulos WHERE (((CodUsu)='
       + inttostr(usu) + ')) AND (((Modulo)=' + inttostr(area) + '));';
  DtsUni.Open;
  if DtsUni.RecordCount = 0 then Result:=4
  else begin
    modulo2.UsuNvl:=DtsUni.Fields[0].AsInteger;
    modulo2.num_modulo:=inttostr(area);
    Result:=DtsUni.Fields[0].AsInteger;
  end;
  DtsUni.Close;
except
  Result:=-1;
end;
DtsUni.Free;
end;

Function CreaPeriodosStrGrid(var Grilla:TStringGrid; ColAnio,ColCuota,ColControl:byte;CharControl:string):widestring;
var j,an:integer;
    cont:widestring;
    kf:boolean;
    rq:string;
begin
{Damos los Periodos involucrados}
cont:='';kf:=false;an:=0;
for j:=1 to Grilla.RowCount-1 do begin
  if (trim(Grilla.Cells[ColControl,j])=CharControl) or (CharControl='') then begin
    if an <> strtoint(trim(Grilla.Cells[ColAnio,j])) then begin
      if trim(cont) <> '' then cont:=cont + '/' + inttostr(an) + ' - '
      else begin
        if an=0 then an:=strtoint(trim(Grilla.Cells[ColAnio,j]))
        else cont:=cont + '/' + inttostr(an) + ' - ';
      end;
      an:=strtoint(trim(Grilla.Cells[ColAnio,j]));
      kf:=false;
    end;
    if not(kf) then begin
      cont:=cont + StrFill(Grilla.Cells[ColCuota,j],2,'0','R');
      kf:=true;
    end
    else cont:=cont + ',' + completa1(Grilla.Cells[ColCuota,j],2,'0');
  end;
end;
cont:=cont + '/' + inttostr(an) + ' - ';
Result:=cont;
end;


Function CreaPeriodosDBGrid(var Grilla:TADODataSet; ColAnio,ColCuota,ColControl:byte;CharControl:string):widestring;
var j,an:integer;
    cont:widestring;
    kf:boolean;
    rq:string;
begin
{Damos los Periodos involucrados}
cont:='';kf:=false;an:=0;
for j:=1 to Grilla.RecordCount do begin
  Grilla.RecNo:=j;
  if (trim(Grilla.Fields[ColControl].AsString)=CharControl) or (CharControl='') then begin
    try
      if an <> StrToInt(trim(Grilla.Fields[ColAnio].AsString)) then begin
        if trim(cont) <> '' then cont:=cont + '/' + inttostr(an) + ' - '
        else begin
          if an=0 then an:=Grilla.Fields[ColAnio].AsInteger
          else cont:=cont + '/' + inttostr(an) + ' - ';
        end;
        an:=Grilla.Fields[ColAnio].AsInteger;
        kf:=false;
      end;
      if not(kf) then begin
        cont:=cont + StrFill(Grilla.Fields[ColCuota].AsString,2,'0','R');
        kf:=true;
      end
      else cont:=cont + ',' + completa1(Grilla.Fields[ColCuota].AsString,2,'0');
    except
      break
    end;
  end;
end;
cont:=cont + '/' + inttostr(an) + ' - ';
Result:=cont;
end;

{Funcion que genera el codigo verificador, recibe un string
de cualquier largo de caracteres y devuelve otro igaul mas uno}
function codigo_ver(cadena:string;EsCodbar:boolean):string;
var ro,total,interme:integer;
    constante,el_codigo:string;
begin
if EsCodBar then begin //Caso para los codigos de barra
  // Calculamos la constante de acuerdo al largo de la cadena
  constante:='13579';
  if Length(cadena) > 5 then begin
    total:=1;
    for ro:=4 to Length(cadena) - 1 do begin
      case total of
        1:constante:=constante + '3';
        2:constante:=constante + '5';
        3:constante:=constante + '7';
        4:begin
          constante:=constante + '9';
          total:=0;
          end
      end;
      total:=total + 1;
    end;
  end;
  total:=0;ro:=0;
  for ro:=1 to Length(cadena)  do begin
    interme:=strtoint(cadena[ro]);
    interme:=interme * strtoint(constante[ro]);
    total:=total + interme;
  end;
  total:=trunc(total/2);
  el_codigo:=IntToStr(Total mod 10);
end
else begin // Caso para el cuim, cuit o cuil
  constante:='5432765432';

  total:=0;ro:=0;interme:=0;
  for ro:=10 downto 1 do begin
    interme:=strtoint(cadena[ro]);
    interme:=interme * strtoint(constante[ro]);
    total:=total + interme;
  end;
  if (total mod 11) = 0 then el_codigo:='0'
  else begin
    if (total mod 11) > 1 then el_Codigo:=IntToStr(11-(Total mod 11))
    else el_codigo:='1'
  end;
end;

Result:=cadena + el_Codigo;
end;


FUNCTION emailValido(CONST Value: String): boolean;
  FUNCTION CheckAllowed(CONST s: String): boolean;
  VAR i: Integer;
  BEGIN
  Result:= False;
  FOR i:= 1 TO Length(s) DO // illegal char in s -> no valid address
  IF NOT (s[i] IN ['a'..'z','A'..'Z','0'..'9','_','-','.']) THEN Exit;
  Result:= true;
  END;
VAR
  i,len: Integer;
  namePart, serverPart: String;
BEGIN // of IsValidEmail
  Result:= False;
  i:= Pos('@', Value);
  IF (i=0) OR (Pos('..',Value) > 0) THEN Exit;
  namePart:= Copy(Value, 1, i - 1);
  serverPart:= Copy(Value,i+1,Length(Value));
  len:=Length(serverPart);
  // must have dot and at least 3 places from end, 2 places from begin
  IF (len<4) OR
     (Pos('.',serverPart)=0) OR
     (serverPart[1]='.') OR
     (serverPart[len]='.') OR
     (serverPart[len-1]='.') THEN Exit;
  Result:= CheckAllowed(namePart) AND CheckAllowed(serverPart);
END;

end.

