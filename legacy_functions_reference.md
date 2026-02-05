# Referencia de Funciones Legacy VB6

Este documento recopila la documentación de funciones legacy VB6 relevantes para la migración, con su lógica resumida y comportamiento.


## Índice

1. [GEN_VALFEC](#gen_valfec)
2. [PMF_ERROR](#pmf_error)
3. [GRABAR_MOVIMIENTO](#grabar_movimiento)
4. [INSERTAR_SUPLEMENTO_TCSUPL](#insertar_suplemento_tcsupl)
   - [UPDATE DTPOLI (Actualización de Póliza)](#update-dtpoli---actualización-de-póliza)
   - [PMS_BORRA_TMPROR (Limpieza de Temporales)](#limpieza-de-temporales-pms_borra_tmpror)
   - [IntegrarConBDI_ODBC (Integración con BDI)](#integración-con-bdi)

---

## GEN_VALFEC

**Módulo:** `GENRT.BAS`

**Propósito:** Valida si una fecha en formato `dd/mm/yyyy` es correcta.

Esto será un parámetro de entrada de tipo date. No hará falta validar si es una fecha o no.

---

## PMF_ERROR

**Módulo:** `SDGEN1.BAS`

**Propósito:** Muestra mensajes de error personalizados obtenidos de la base de datos y retorna el código de la tecla pulsada por el usuario.


### Parámetros

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `errorCode` | String | Código del error a buscar (ej: "GEI024") |


### Query lanzada
```sql
SELECT 
    MERRCDER as errorCode, 
    MERRDSER as errorDescription, 
    MERRCDTE as errorTypeCode, 
    MERRDSCT as shortDescription 
FROM DTMERR 
WHERE
    MERRCDER = :errorCode
```
---

## GRABAR_MOVIMIENTO

**Módulo:** `IBERCAJA.BAS`

**Propósito:** Inserta un registro en la tabla `TCMOVI` para auditoría de movimientos de pólizas. Registra cambios realizados en las pólizas con valores antes/después y metadatos del movimiento.

### Firma

```vb
Function GRABAR_MOVIMIENTO(sTipo As String, _
                           sDelegacion As String, _
                           sFecha As String, _
                           sPoliza As String, _
                           sCertificado As String, _
                           sUsuario As String, _
                           sSuplemento As String, _
                           sCliente As String, _
                           sAntes As String, _
                           sDespues As String, _
                           Optional sFecha_Efecto As String, _
                           Optional sNumExpSGO As String) As Integer
```

### Parámetros

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `movementType` | String | Código del tipo de movimiento (ver constantes G_MOV_*) |
| `delegationCode` | String | Código de delegación (635=Salud, 634=Decesos) |
| `policyNumber` | Integer | Número de póliza |
| `certificateCode` | Integer | Número de certificado |
| `movementDate` | LocalDate | Fecha del movimiento en formato YYYYMMDD |
| `user` | String | Usuario que realiza el movimiento |
| `supplementNumber` | String | Número de suplemento |
| `clientCode` | String | Código de cliente (puede ser NULL)
| `beforeValue` | String | Valor antes del cambio |
| `afterValue` | String | Valor después del cambio |
| `modifiedAt` | String (Opcional) | Fecha de efecto del movimiento en formato YYYYMMDD |
| `sgoTask` | Long (Opcional) | Número de expediente SGO |


### Constantes de Tipos de Movimiento (G_MOV_*)

```vb
' Constantes definidas en IBERCAJA.BAS
Public Const G_MOV_ALTA_POLIZA As String = "01"
Public Const G_MOV_BAJA_POLIZA As String = "02"
Public Const G_MOV_MODIFICACION_POLIZA As String = "03"
Public Const G_MOV_SUPLEMENTO As String = "04"
Public Const G_MOV_ANULACION As String = "05"
Public Const G_MOV_REHABILITACION As String = "06"
Public Const G_MOV_CAMBIO_TOMADOR As String = "07"
Public Const G_MOV_CAMBIO_AGENTE As String = "08"
Public Const G_MOV_CAMBIO_DOMICILIO As String = "09"
Public Const G_MOV_CAMBIO_FORMA_PAGO As String = "10"
Public Const G_MOV_CAMBIO_CUENTA As String = "11"
Public Const G_MOV_ALTA_ASEGURADO As String = "12"
Public Const G_MOV_BAJA_ASEGURADO As String = "13"
Public Const G_MOV_MODIFICACION_ASEGURADO As String = "14"
Public Const G_MOV_CAMBIO_COBERTURA As String = "15"
Public Const G_MOV_CAMBIO_TARIFA As String = "16"
Public Const G_MOV_CAMBIO_DESCUENTO As String = "17"
Public Const G_MOV_RENOVACION As String = "18"
Public Const G_MOV_EMISION_RECIBO As String = "19"
Public Const G_MOV_ANULACION_RECIBO As String = "20"
```

### SQL Generada Dinámicamente

La función construye la SQL de forma **DINÁMICA**, añadiendo campos opcionales solo si tienen valor:

#### Versión MÍNIMA (solo campos obligatorios)

```sql
INSERT INTO TCMOVI(
    MOVINUME, 
    MOVITIPO, 
    MOVICDDE, 
    MOVIFECH, 
    MOVINPOL, 
    MOVICDCE, 
    MOVIUSUA, 
    MOVINUSU, 
    MOVICDCL) 
VALUES (
    SQ_MOVIM.NEXTVAL,
    :movementType,
    :delegationCode,
    :movementDate,
    :policyNumber,
    :certificateCode,
    :user,
    :supplementNumber,
    :clientCode)
```

#### Versión COMPLETA (todos los campos opcionales informados)

Cuando todos los campos opcionales tienen valor:

```sql
INSERT INTO TCMOVI(
    MOVINUME,
    MOVITIPO,
    MOVICDDE,
    MOVINPOL,
    MOVICDCE,
    MOVIFECH,
    MOVIUSUA,
    MOVINUSU,
    MOVICDCL,
    MOVIANTE,
    MOVIDESP,
    MOVIFECH_EFECTO,
    MOVITAREA_SGO) 
VALUES (
    SQ_MOVIM.NEXTVAL,
    :movementType,
    :delegationCode,
    :policyNumber,
    :certificateCode,
    :movementDate,
    :user,
    :supplementNumber,
    :clientCode,
    :beforeValue,
    :afterValue,
    :modifiedAt,
    :sgoTask)
```

### Estructura de Base de Datos

**Tabla:** `TCMOVI`

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `MOVINUME` | NUMBER(9,0) | Secuencia |
| `MOVITIPO` | VARCHAR2(2) | Tipo de movimiento (G_MOV_*) |
| `MOVICDDE` | VARCHAR2(3) | Código de delegación |
| `MOVINPOL` | NUMBER(9,0) | Número de póliza |
| `MOVICDCE` | NUMBER(9,0) | Código de certificado |
| `MOVIFECH` | VARCHAR2(8) | Fecha del movimiento (YYYYMMDD) |
| `MOVIUSUA` | VARCHAR2(10) | Usuario que realiza el movimiento |
| `MOVINUSU` | VARCHAR2(3) | Número de suplemento |
| `MOVICDCL` | VARCHAR2(10) | Código de cliente |
| `MOVIANTE` | VARCHAR2(200) | Valor antes del cambio |
| `MOVIDESP` | VARCHAR2(200) | Valor después del cambio |
| `MOVIFECH_EFECTO` | VARCHAR2(8) | Fecha de efecto (YYYYMMDD) |
| `MOVITAREA_SGO` | VARCHAR2(15) | Número de expediente SGO |
| `MOVI_REGU_ANTICIPO` | VARCHAR2(1) | Marca si se debe insertar un apunte para regularizar el anticipo (S/N) |

### Código VB6 Reorganizado.

```vb
Public Sub GRABAR_MOVIMIENTO(
    sTipo As String,
    sDelegacion As String,
    sFecha As String,
    sPoliza As String,
    sCertificado As String,
    sUsuario As String,
    sSuplemento As String,
    sCliente As String,
    sAntes As String,
    sDespues As String,
    Optional sFecha_Efecto = "",
    Optional sNumExpSGO As String = "")
    'Procedimiento que inserta un registro en la tabla TCMOVI.

    Dim Stat As String
    Dim sql_Fields As String
    Dim sql_Values As String
    
    Dim HSTMT As Long
    Dim ENDCODE As Integer
    
    sql_Fields = ""
    sql_Fields = "MOVINUME,"
    sql_Fields = sql_Fields & " MOVITIPO,"
    sql_Fields = sql_Fields & " MOVICDDE,"
    sql_Fields = sql_Fields & " MOVIFECH,"
    sql_Fields = sql_Fields & " MOVINPOL,"
    sql_Fields = sql_Fields & " MOVICDCE,"
    sql_Fields = sql_Fields & " MOVIUSUA,"
    sql_Fields = sql_Fields & " MOVINUSU,"
    sql_Fields = sql_Fields & " MOVICDCL"
    
    sql_Values = ""
    sql_Values = "SQ_MOVIM.NEXTVAL,"
    sql_Values = sql_Values & "'" & Trim(sTipo) & "',"
    sql_Values = sql_Values & Trim(sDelegacion) & ","
    sql_Values = sql_Values & "'" & Trim(sFecha) & "',"
    sql_Values = sql_Values & Trim(sPoliza) & ","
    sql_Values = sql_Values & Trim(sCertificado) & ","
    sql_Values = sql_Values & "'" & Trim(sUsuario) & "',"
    sql_Values = sql_Values & Trim(sSuplemento) & ","

    If Trim(sCliente) = "" Then
      sql_Values = sql_Values & "NULL"
    Else
      sql_Values = sql_Values & Trim(sCliente)
    End If
    
    If UCase(sAntes) <> "NULL" And NoNulls(sAntes) <> "" Then
        sql_Fields = sql_Fields & ", MOVIANTE"
        sql_Values = sql_Values & ",'" & Trim(sAntes) & "'"
    End If
    
    If UCase(sDespues) <> "NULL" And NoNulls(sDespues) <> "" Then
        sql_Fields = sql_Fields & ", MOVIDESP"
        sql_Values = sql_Values & ",'" & Trim(sDespues) & "'"
    End If
    
    If NoNulls(sFecha_Efecto) <> "" Then
        sql_Fields = sql_Fields & ", MOVIFECH_EFECTO"
        sql_Values = sql_Values & ",'" & Trim(sFecha_Efecto) & "'"
    End If
    
    If NoNulls(sNumExpSGO) <> "" Then
        sql_Fields = sql_Fields & ", MOVITAREA_SGO"
        sql_Values = sql_Values & ",'" & Trim(sNumExpSGO) & "'"
    End If
    
    Stat = "INSERT INTO TCMOVI(" & sql_Fields & ") VALUES (" & sql_Values & ")"
    
    HSTMT = SQL_EXEC(G_HDBC, Stat, 0)
    ENDCODE = SQL_END(HSTMT)
End Sub
```

### Lógica Resumida

1. **Manejo de valores opcionales**: Añade campos a la SQL dinámica según si los valores opcionales si tienen información (que no sea nula o vacía)
2. **Inserción**: Ejecuta INSERT en `TCMOVI`

### Notas de Migración

1. **Formato de fechas**: La base de datos usa formato `YYYYMMDD` como String. Es necesario utilizar funciones de conversión.
2. **Delegaciones**: En TCMOVI existen 2 códigos de delegación (635=Salud, 634=Decesos).
3. **Transaccionalidad**: El movimiento de auditoría debería grabarse en la misma transacción que la operación principal para garantizar consistencia. Considerar usar `@Transactional(propagation = Propagation.MANDATORY)` si debe ejecutarse dentro de una transacción existente.
4. **Seguridad**: El campo `sUsuario` viene del contexto de sesión VB6. En Spring, obtener del `SecurityContext` o inyectar mediante AOP.
5. **Código existente en el proyecto**: Verificar si `MovementEntity` y `MovementService` existentes en el proyecto pueden extenderse o si se requiere una nueva entidad específica para auditoría.
---

## INSERTAR_SUPLEMENTO_TCSUPL

**Módulo:** `mdlSuplementos.bas`

**Propósito:** Inserta un registro de suplemento en la tabla TCSUPL con todos los datos históricos de una póliza/certificado. Construye dinámicamente el INSERT según si hay datos de domicilio nuevos o antiguos. Delega en INSERTAR_SUPLEMENTO_TSSUPC para las cláusulas asociadas.


### SQL Generada Dinámicamente

La función construye la SQL de forma **DINÁMICA**, con dos variantes según el formato del domicilio:

#### Versión MÍNIMA (domicilio en formato antiguo - sin "#")

Cuando `G_POCE2DOMI` **NO contiene** el carácter `#`, se considera formato antiguo y solo se graba `SUP_NOMBREVIA`:

```sql
INSERT INTO TCSUPL (
   -- Campos de identificación (4 campos)
  SUPLCDDE, SUPLNPOL, SUPLCDCE, SUPLNUSU,
  
  -- Campos de tipo y estado (11 campos)
  SUPLTIPO, SUPLSITP, SUPLSITC, SUPLFECA, SUPLFECB, SUPLFECC,
  SUPLFEBA, SUPLCDPT, SUPLCDTA, SUPLFOPA, SUPLTIPA,
  
  -- Campos de personas y referencias (9 campos)
  SUPLNUPE, SUPLIDCP, SUPLIDCO, SUPLCDTR, SUPLIDEX,
  SUPLCOBR, SUPLAGTA, SUPLAGTB, SUPLINSP,
  
  -- Campos de primas e importes (10 campos)
  SUPLPRNT, SUPLPRNE, SUPLRECA, SUPLRECE, SUPLIMPT, SUPLIMPE,
  SUPLTORE, SUPLMOCE, SUPLCDPS, SUPLCDPO,
  
  -- Campos adicionales de importes (14 campos)
  SUPLTFNO_NUTE, SUPLCRNT, SUPLCRNE, SUPLCECA, SUPLCECE,
  SUPLCMPT, SUPLCMPE, SUPLCORE, SUPLIPUN, SUPLIPUE,
  SUPLCPUN, SUPLCPUE, SUPLIDMA, SUPLNUCE,
  
  -- Campos de referencia (2 campos)
  SUPLCDRP, SUPLCDPC,
  
  -- Campos P207 - provincia y switches (7 campos)
  SUPLPROVINCIA_TARIFICACION, SUPLSWPROVINCIA, SUPLSWPRODUCCION,
  SUPLSWTARIFA, SUPLSWDCTO, SUPLDCTONUMPERSONAS, SUPLRECFORMAPAGO,
  
  -- Campos de adaptación (2 campos)
  SUPLADAP, SUPLTERRITORIALIDAD,
  
  -- Campos de domicilio (1 campo, formato viejo)
  SUP_NOMBREVIA
  
  -- Campos de descuento - 8 grupos
  SUPLDESC_G01, SUPLDESC_G02, SUPLDESC_G03, SUPLDESC_G04,
  SUPLDESC_G05, SUPLDESC_G06, SUPLDESC_G07, SUPLDESC_G08,
  
  -- Indicadores (2 campos)
  SUPLINDVTAC, SUPLINDPRIC
) VALUES (
   -- Campos de identificación (4 campos)
   :delegationCode, :policyNumber, :certificateNUmber, :supplementNmber,

   -- Campos de tipo y estado (11 campos)
   '03', 'A', 'A', '20241115', '20240101', '20240101',
   NULL, 100, 'T01', 'M', 'D',

   -- Campos de personas y referencias (9 campos)
   2, 'X1234567A', 'Y9876543B', '01', 'S',
   'S', 'AG001', 'AG002', 'S',

   -- Campos de primas e importes (10 campos)
   1200.50, 0, 120.05, 0, 252.11, 0,
   1572.66, 'S', '28001', '28',

   -- Campos adicionales de importes (14 campos)
   '912345678', 0, 0, 0, 0,
   0, 0, 0, 0, 0,
   0, 0, NULL, 0, 

   -- Campos de referencia (2 campos)
   0, 0,

   -- Campos P207 - provincia y switches (7 campos)
   '28', 'S', 'S',
   'N', 'N', 0, 0,

   -- Campos de adaptación (2 campos)
   'A', 'N',

   -- Campos de domicilio (1 campo, formato viejo)
   'CALLE GRAN VIA 123, 5º B',

   -- Campos de descuento - 8 grupos
   10.00, 5.00, 0, 0, 0, 0, 0, 0,
   
   -- Indicadores (2 campos)
   'S', 'N'
)
```

#### Versión COMPLETA (domicilio normalizado - con "#")

Cuando `G_POCE2DOMI` **contiene** el carácter `#`, se parsean los 11 campos de domicilio normalizado:

```sql
INSERT INTO TCSUPL (
   -- Campos de identificación (4 campos)
  SUPLCDDE, SUPLNPOL, SUPLCDCE, SUPLNUSU,
  
  -- Campos de tipo y estado (11 campos)
  SUPLTIPO, SUPLSITP, SUPLSITC, SUPLFECA, SUPLFECB, SUPLFECC,
  SUPLFEBA, SUPLCDPT, SUPLCDTA, SUPLFOPA, SUPLTIPA,
  
  -- Campos de personas y referencias (9 campos)
  SUPLNUPE, SUPLIDCP, SUPLIDCO, SUPLCDTR, SUPLIDEX,
  SUPLCOBR, SUPLAGTA, SUPLAGTB, SUPLINSP,
  
  -- Campos de primas e importes (10 campos)
  SUPLPRNT, SUPLPRNE, SUPLRECA, SUPLRECE, SUPLIMPT, SUPLIMPE,
  SUPLTORE, SUPLMOCE, SUPLCDPS, SUPLCDPO,
  
  -- Campos adicionales de importes (14 campos)
  SUPLTFNO_NUTE, SUPLCRNT, SUPLCRNE, SUPLCECA, SUPLCECE,
  SUPLCMPT, SUPLCMPE, SUPLCORE, SUPLIPUN, SUPLIPUE,
  SUPLCPUN, SUPLCPUE, SUPLIDMA, SUPLNUCE,
  
  -- Campos de referencia (2 campos)
  SUPLCDRP, SUPLCDPC,
  
  -- Campos P207 - provincia y switches (7 campos)
  SUPLPROVINCIA_TARIFICACION, SUPLSWPROVINCIA, SUPLSWPRODUCCION,
  SUPLSWTARIFA, SUPLSWDCTO, SUPLDCTONUMPERSONAS, SUPLRECFORMAPAGO,
  
  -- Campos de adaptación (2 campos)
  SUPLADAP, SUPLTERRITORIALIDAD,
  
  -- Campos de domicilio (11 campos, formato nuevo)
  SUP_CDG_TIPOVIA, SUP_NOMBREVIA, SUP_NUMEROVIA,
  SUP_PORTAL, SUP_BLOQUE, SUP_ESCALERA,
  SUP_PISO, SUP_PUERTA, SUP_RESTOVIA,
  SUP_CPOBLA_INE, SUP_CVIA_INE,
  
  -- Campos de descuento - 8 grupos
  SUPLDESC_G01, SUPLDESC_G02, SUPLDESC_G03, SUPLDESC_G04,
  SUPLDESC_G05, SUPLDESC_G06, SUPLDESC_G07, SUPLDESC_G08,
  
  -- Indicadores (2 campos)
  SUPLINDVTAC, SUPLINDPRIC
) VALUES (
   -- Campos de identificación (4 campos)
   :delegationCode, :policyNumber, :certificateNUmber, :supplementNmber,

   -- Campos de tipo y estado (11 campos)
   '03', 'A', 'A', '20241115', '20240101', '20240101',
   NULL, 100, 'T01', 'M', 'D',

   -- Campos de personas y referencias (9 campos)
   2, 'X1234567A', 'Y9876543B', '01', 'S',
   'S', 'AG001', 'AG002', 'S',

   -- Campos de primas e importes (10 campos)
   1200.50, 0, 120.05, 0, 252.11, 0,
   1572.66, 'S', '28001', '28',

   -- Campos adicionales de importes (14 campos)
   '912345678', 0, 0, 0, 0,
   0, 0, 0, 0, 0,
   0, 0, NULL, 0, 

   -- Campos de referencia (2 campos)
   0, 0,

   -- Campos P207 - provincia y switches (7 campos)
   '28', 'S', 'S',
   'N', 'N', 0, 0,

   -- Campos de adaptación (2 campos)
   'A', 'N',

   -- Campos de domicilio (11 campos)
   'CL',           -- SUP_CDG_TIPOVIA (código tipo vía)
   'GRAN VIA',     -- SUP_NOMBREVIA
   '123',          -- SUP_NUMEROVIA
   '',             -- SUP_PORTAL
   '',             -- SUP_BLOQUE
   'A',            -- SUP_ESCALERA
   '5',            -- SUP_PISO
   'B',            -- SUP_PUERTA
   '',             -- SUP_RESTOVIA (resto dirección)
   '2807901',      -- SUP_CPOBLA_INE (código población INE)
   '280790100123', -- SUP_CVIA_INE (código vía INE)

   -- Campos de descuento - 8 grupos
   10.00, 5.00, 0, 0, 0, 0, 0, 0,

   -- Indicadores (2 campos)
   'S', 'N'
)
```

### Firma

```vb
Public Sub INSERTAR_SUPLEMENTO_TCSUPL(
    ByVal G_POCE2CDDE As String,
    ByVal G_POCE2NPOL As String,
    ByVal G_POCE2CDCE As String,
    ByVal G_POCE2NUSU As String,
    ByVal G_CER As String,
    ByVal G_POLI2ESPA As String,
    ByVal G_POCE2ALBA As String,
    ByVal G_POCE2FECB As String,
    ByVal G_POCE2FECA As String,
    ByVal G_POCE2FEBA As String,
    ByVal G_POLI2CDPT As String,
    ByVal G_POCE2CDTA As String,
    ByVal G_POCE2FOPA As String,
    ByVal G_POCE2TIPA As String,
    ByVal G_POCE2NUPE As String,
    ByVal G_POLI2IDCP As String,
    ByVal G_POLI2IDCO As String,
    ByVal G_POLI2CDTR As String,
    ByVal G_POLI2IDEX As String,
    ByVal G_POCE2COBR As String,
    ByVal G_POCE2AGTA As String,
    ByVal G_POCE2AGTB As String,
    ByVal G_CAT_TIPO_ANT As String,
    ByVal G_POCE2PRNT As String,
    ByVal G_POCE2PRNE As String,
    ByVal G_POCE2RECA As String,
    ByVal G_POCE2RECE As String,
    ByVal G_POCE2IMPT As String,
    ByVal G_POCE2IMPE As String,
    ByVal G_POCE2TORE As String,
    ByVal G_POCE2MOBA As String,
    ByVal G_POCE2DOMI As String,
    ByVal G_POCE2CDPS As String,
    ByVal G_POCE2CDPO As String,
    ByVal G_POCE2TFNO As String,
    ByVal G_POCE2CRNT As String,
    ByVal G_POCE2CRNE As String,
    ByVal G_POCE2CECA As String,
    ByVal G_POCE2CECE As String,
    ByVal G_POCE2CMPT As String,
    ByVal G_POCE2CMPE As String,
    ByVal G_POCE2CORE As String,
    ByVal G_POCE2IPUN As String,
    ByVal G_POCE2IPUE As String,
    ByVal G_POCE2CPUN As String,
    ByVal G_POCE2CPUE As String,
    ByVal G_Fecha_Asegurado As String,
    ByVal G_POLI2IDMA As String,
    ByVal G_POLI2NUCE As String,
    ByVal G_POLI2CDRP As String,
    ByRef G_POCEFECM As String,
    ByRef strSuplemento_DTSUAS As String,
    ByVal G_POCEPROVINCIA_TARIFICACION As String,
    ByVal G_POCESWPROVINCIA As String,
    ByVal G_POCESWPRODUCCION As String,
    ByVal G_POCESWTARIFA As String,
    ByVal G_POCESWDCTO As String,
    ByVal G_POCEDCTONUMPERSONAS As String,
    ByVal G_POCERECFORMAPAGO As String,
    G_Descuentos() As SubDescuento)
```

### Parámetros Principales

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `G_POCE2CDDE` | String | Código de delegación |
| `G_POCE2NPOL` | String | Número de póliza |
| `G_POCE2CDCE` | String | Código de certificado (0 para póliza individual) |
| `G_POCE2NUSU` | String | Número de suplemento |
| `G_CER` | String | Tipo de entidad: "01"=Individual, "02"=Colectiva, "03"=Certificado |
| `G_POLI2ESPA` | String | Estado previo de la póliza |
| `G_POCE2ALBA` | String | Estado actual (alta/baja) |
| `G_POCE2FECB` | String | Fecha de efecto B (dd/mm/yyyy) |
| `G_POCE2FECA` | String | Fecha de efecto A (dd/mm/yyyy) |
| `G_POCE2FEBA` | String | Fecha de baja (dd/mm/yyyy) |
| `G_POLI2CDPT` | String | Código de producto |
| `G_POCE2CDTA` | String | Código de tarifa |
| `G_POCE2FOPA` | String | Forma de pago |
| `G_POCE2TIPA` | String | Tipo de pago |
| `G_POCE2NUPE` | String | Número de personas/asegurados |
| `G_POCE2PRNT` | String | Prima neta total |
| `G_POCE2PRNE` | String | Prima neta extranjero |
| `G_POCE2IMPT` | String | Impuestos totales |
| `G_POCE2TORE` | String | Total recibo anual |
| `G_POCE2DOMI` | String | Domicilio (formato antiguo o nuevos campos separados por #) |
| `G_POCEPROVINCIA_TARIFICACION` | String | Provincia para tarificación |
| `G_Descuentos()` | SubDescuento | Array con 8 grupos de descuento (G01-G08) |
| `G_POCEFECM` | String (ByRef) | Fecha de modificación (retornado) |
| `strSuplemento_DTSUAS` | String (ByRef) | Número de suplemento generado (retornado) |


---
Dado que la tabla es enorme, se ha generado una entidad `TCSUPL`
que refleja la tabla entera
#### **Entidad JPA:** `TCSUPL`
```java
package es.caser.salud.polind.entity;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.IdClass;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.math.BigDecimal;

@Entity
@Table(name = "TCSUPL", schema = "SCOTT")
@IdClass(TcsuplPK.class)
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class TcsuplEntity {

    @Id
    @Column(name = "SUPLCDDE")
    private Integer delegationCode;

    @Id
    @Column(name = "SUPLNPOL")
    private Integer policyNumber;

    @Id
    @Column(name = "SUPLCDCE")
    private Integer certificateCode;

    @Id
    @Column(name = "SUPLNUSU")
    private Integer supplementNumber;

    @Column(name = "SUPLTIPO", length = 2)
    private String policyType;

    @Column(name = "SUPLSITP", length = 2)
    private String policyStatus;

    @Column(name = "SUPLSITC", length = 2)
    private String certifcateStatus;

    @Column(name = "SUPLFECA", length = 8)
    private String supplementStartDate;

    @Column(name = "SUPLFECB", length = 8)
    private String supplementEndDate;

    @Column(name = "SUPLFECC", length = 8)
    private String policyStartDate;

    @Column(name = "SUPLFEBA", length = 8)
    private String policyEndDate;

    @Column(name = "SUPLCDPC", length = 2)
    private String collectivePaymentTypeCode;

    @Column(name = "SUPLCDPT")
    private Integer productNumber;

    @Column(name = "SUPLCDTA", length = 5)
    private String rateCode;

    @Column(name = "SUPLFOPA", length = 2)
    private String paymentMethod;

    @Column(name = "SUPLTIPA", length = 2)
    private String paymentTypeCode;

    @Column(name = "SUPLNUPE")
    private Integer insuredNumber;

    @Column(name = "SUPLIDCP", length = 1)
    private String specialPolicyCommissionFlag;

    @Column(name = "SUPLIDCO", length = 1)
    private String specialCoveragesFlag;

    @Column(name = "SUPLCDTR", length = 2)
    private String collectiveTarificationCode;

    @Column(name = "SUPLIDEX", length = 1)
    private String foreignFlag;

    @Column(name = "SUPLCOBR", length = 9)
    private String collector;

    @Column(name = "SUPLAGTA", length = 9)
    private String mainAgent;

    @Column(name = "SUPLAGTB", length = 9)
    private String secondaryAgent;

    @Column(name = "SUPLINSP", length = 9)
    private String inspector;

    @Column(name = "SUPLPRNT", precision = 14, scale = 4)
    private BigDecimal netPremium;

    @Column(name = "SUPLPRNE", precision = 14, scale = 4)
    private BigDecimal foreignNetPremium;

    @Column(name = "SUPLRECA", precision = 14, scale = 4)
    private BigDecimal surcharge;

    @Column(name = "SUPLRECE", precision = 14, scale = 4)
    private BigDecimal foreignSurcharge;

    @Column(name = "SUPLIMPT", precision = 14, scale = 4)
    private BigDecimal taxRate;

    @Column(name = "SUPLIMPE", precision = 14, scale = 4)
    private BigDecimal foreignTaxRate;

    @Column(name = "SUPLTORE", precision = 14, scale = 4)
    private BigDecimal totalReceipt;

    @Column(name = "SUPLMOBP", length = 2)
    private String policyCancellationReason;

    @Column(name = "SUPLMOCE", length = 2)
    private String certificateCancellationReason;

    @Column(name = "SUPLIDMA", length = 1)
    private String autoAssignCertificateNumbersFlag;

    @Column(name = "SUPLDOMI", length = 80)
    private String benefitAddress;

    @Column(name = "SUPLCDPS", length = 5)
    private String benefitPostalCode;

    @Column(name = "SUPLCDPO", length = 7)
    private String suplcdpo;

    @Column(name = "SUPLTFNO")
    private Integer supltfno;

    @Column(name = "SUPLNUCE")
    private Integer suplnuce;

    @Column(name = "SUPLCRNT", precision = 14, scale = 4)
    private BigDecimal suplcrnt;

    @Column(name = "SUPLCRNE", precision = 14, scale = 4)
    private BigDecimal suplcrne;    

    @Column(name = "SUPLCECA", precision = 14, scale = 4)
    private BigDecimal suplceca;

    @Column(name = "SUPLCECE", precision = 14, scale = 4)
    private BigDecimal suplcece;

    @Column(name = "SUPLCMPT", precision = 14, scale = 4)
    private BigDecimal suplcmpt;

    @Column(name = "SUPLCMPE", precision = 14, scale = 4)
    private BigDecimal suplcmpe;

    @Column(name = "SUPLCORE", precision = 14, scale = 4)
    private BigDecimal suplcore;

    @Column(name = "SUPLIPUN", precision = 14, scale = 4)
    private BigDecimal suplipun;

    @Column(name = "SUPLIPUE", precision = 14, scale = 4)
    private BigDecimal suplipue;

    @Column(name = "SUPLCPUN", precision = 14, scale = 4)
    private BigDecimal suplcpun;

    @Column(name = "SUPLCPUE", precision = 14, scale = 4)
    private BigDecimal suplcpue;

    @Column(name = "SUPLICON", precision = 14, scale = 4)
    private BigDecimal suplicon;

    @Column(name = "SUPLCCON", precision = 14, scale = 4)
    private BigDecimal suplccon;

    @Column(name = "SUPLPRDE", precision = 14, scale = 4)
    private BigDecimal suplprde;

    @Column(name = "SUPLCRDE", precision = 14, scale = 4)
    private BigDecimal suplcrde;

    @Column(name = "SUPLBENE", length = 40)
    private String suplbene;

    @Column(name = "SUPLCDPA", length = 2)
    private String countryCode;

    @Column(name = "SUPLCDRP")
    private Integer paymentRuleCode;

    @Column(name = "SUPLPROVINCIA_TARIFICACION", length = 2)
    private String tarificationProvincia;

    @Column(name = "SUPLSWPROVINCIA", length = 1)
    private String updatePolicyFlag = "S";

    @Column(name = "SUPLSWTARIFA", length = 1, nullable = false)
    private String updateRateFlag = "N";

    @Column(name = "SUPLSWDCTO", length = 1, nullable = false)
    private String updateDiscountFlag = "N";

    @Column(name = "SUPLSWPRODUCCION", length = 1, nullable = false)
    private String useProductionRateFlag = "N";

    @Column(name = "SUPLTIPO_DCTO", length = 3, nullable = false)
    private String certificateDiscountType = "000";

    @Column(name = "SUPLDCTONUMPERSONAS", precision = 7, scale = 4)
    private BigDecimal peopleNumberDiscount = 0;

    @Column(name = "SUPLRECFORMAPAGO", precision = 7, scale = 4)
    private BigDecimal paymentMethodSurchage = 0;

    @Column(name = "SUPLADAP", length = 2)
    private String adaptedPolicy;

    @Column(name = "SUPLTERRITORIALIDAD", length = 2)
    private String territoriality;

    @Column(name = "SUP_CDG_TIPOVIA", length = 10)
    private String streetType;

    @Column(name = "SUP_NOMBREVIA", length = 100)
    private String streetName;

    @Column(name = "SUP_NUMEROVIA", length = 5)
    private String streetNumber;

    @Column(name = "SUP_PORTAL", length = 3)
    private String streetGate;

    @Column(name = "SUP_BLOQUE", length = 3)
    private String blockNumber;

    @Column(name = "SUP_ESCALERA", length = 3)
    private String stairsNumber;

    @Column(name = "SUP_PISO", length = 4)
    private String floor;

    @Column(name = "SUP_PUERTA", length = 5)
    private String door;

    @Column(name = "SUP_RESTOVIA", length = 100)
    private String otherStreetInfo;

    @Column(name = "SUP_COD_BDI", length = 10)
    private String bdiCertificateCode;

    @Column(name = "SUP_CPOBLA_INE", length = 20)
    private String ineTownCode;

    @Column(name = "SUP_CVIA_INE", length = 20)
    private String ineStreetCode;

    @Column(name = "SUPLDESC_G01", precision = 7, scale = 4)
    private BigDecimal discountGroup1Value;

    @Column(name = "SUPLDESC_G02", precision = 7, scale = 4)
    private BigDecimal discountGroup2Value;

    @Column(name = "SUPLDESC_G03", precision = 7, scale = 4)
    private BigDecimal discountGroup3Value;

    @Column(name = "SUPLDESC_G04", precision = 7, scale = 4)
    private BigDecimal discountGroup4Value;

    @Column(name = "SUPLDESC_G05", precision = 7, scale = 4)
    private BigDecimal discountGroup5Value;

    @Column(name = "SUPLDESC_G06", precision = 7, scale = 4)
    private BigDecimal discountGroup6Value;

    @Column(name = "SUPLDESC_G07", precision = 7, scale = 4)
    private BigDecimal discountGroup7Value;

    @Column(name = "SUPLDESC_G08", precision = 7, scale = 4)
    private BigDecimal discountGroup8Value;

    @Column(name = "SUPLINDVTAC", length = 1)
    private String policyCrossSellingFlag;

    @Column(name = "SUPLINDPRIC", length = 1)
    private String policyPricingFlag;

    @Column(name = "SUPLTFNO_NUTE", length = 15)
    private String telephoneNumber;

    @Column(name = "SUPL_ID_ANONIM")
    private Integer policyAnonymizationId;
}

package es.caser.salud.polind.entity;

import jakarta.persistence.Column;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.io.Serializable;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode
public class TcsuplPK implements Serializable {

    @Column(name = "SUPLCDDE")
    private Integer delegationCode;

    @Column(name = "SUPLNPOL")
    private Integer policyNumber;

    @Column(name = "SUPLCDCE")
    private Integer certificateCode;

    @Column(name = "SUPLNUSU")
    private Integer supplementNumber;
}
```

**Tabla secundaria:** `TSSUPC` (delegada a INSERTAR_SUPLEMENTO_TSSUPC)

### Código VB6 Original (Fragmento Principal)

```vb
Public Sub INSERTAR_SUPLEMENTO_TCSUPL(...)
   Dim strSql As String
   Dim ihstmt_DTSUPL As Long
   Dim iEnd As Integer
   
   Screen.MousePointer = 11
   
   If Trim$(G_Fecha_Asegurado) <> "" Then
      G_POCEFECM = G_Fecha_Asegurado
   Else
      G_POCEFECM = Format(Now, "dd/mm/yyyy")
   End If
    
   strSql = "INSERT INTO TCSUPL (SUPLCDDE, SUPLNPOL, SUPLCDCE, SUPLNUSU,"
   strSql = strSql & " SUPLTIPO, SUPLSITP, SUPLSITC, SUPLFECA, SUPLFECB, SUPLFECC,"
   strSql = strSql & " SUPLFEBA, SUPLCDPT, SUPLCDTA, SUPLFOPA, SUPLTIPA,"
   ' ... (continúa construcción dinámica del INSERT)
   
   ' Lógica de domicilio: si contiene "#" parsea campos nuevos, sino usa formato antiguo
   If InStr(G_POCE2DOMI, "#") = 0 Then
        strSql = strSql & " SUP_NOMBREVIA"
   Else
        strSql = strSql & " SUP_CDG_TIPOVIA, SUP_NOMBREVIA, SUP_NUMEROVIA,"
        ' ... 11 campos de domicilio normalizado
   End If
   
   ' Añade campos de descuento G01-G08
   strSql = strSql & ", SUPLDESC_G01, SUPLDESC_G02, ... , SUPLDESC_G08"
   strSql = strSql & ", SUPLINDVTAC, SUPLINDPRIC"
   
   ihstmt_DTSUPL = SQL_EXEC(G_HDBC, strSql, 0)
   iEnd = SQL_END(ihstmt_DTSUPL)
   
   ' Delegar inserción de cláusulas si no es póliza colectiva
   If G_CER <> "02" Then
        INSERTAR_SUPLEMENTO_TSSUPC G_POCE2NPOL, G_POCE2CDCE, G_POCE2NUSU
   End If
   
   strSuplemento_DTSUAS = G_POCE2NUSU
   Screen.MousePointer = 0
End Sub
```

### Lógica Resumida

1. **Establecer fecha de modificación**: Si se proporciona `G_Fecha_Asegurado`, usa esa fecha; sino, usa la fecha actual
2. **Construir INSERT dinámico**: 
   - Cabecera fija con ~60 columnas principales
   - Si `G_POCE2DOMI` NO contiene "#": solo graba `SUP_NOMBREVIA` (formato antiguo)
   - Si `G_POCE2DOMI` contiene "#": parsea y graba los 11 campos de domicilio normalizado
3. **Añadir campos de descuento**: Añade los 8 grupos de descuento (G01-G08) desde el array
4. **Añadir indicadores**: `SUPLINDVTAC` y `SUPLINDPRIC` desde variables globales del módulo
5. **Ejecutar INSERT**: Mediante `SQL_EXEC`
6. **Delegar cláusulas**: Si NO es póliza colectiva (`G_CER <> "02"`), llama a `INSERTAR_SUPLEMENTO_TSSUPC` para insertar las cláusulas de la matriz `MAT_TSPOPC`
7. **Retornar número de suplemento**: En `strSuplemento_DTSUAS`

### Variables Globales del Módulo Utilizadas

| Variable | Uso |
|----------|-----|
| `G_HDBC` | Handle de conexión a base de datos |
| `MAT_TSPOPC` | Matriz con datos de cláusulas cargada previamente |
| `G_SUPLADAP_SUP` | Valor de adaptación del suplemento |
| `G_SUPLTERRITORIALIDAD_SUP` | Valor de territorialidad |
| `G_SUPLINDVTAC_SUP` | Indicador de venta cruzada |
| `G_SUPLINDPRIC_SUP` | Indicador de pricing |

### Funciones Auxiliares Utilizadas

- `Grabar_BBDD(valor, tipo)`: Formatea valor para SQL ("N"=numérico, "S"=string)
- `Grabacion_Importe(valor, flag)`: Formatea importes decimales
- `GEN_DTOC1(fecha)`: Convierte fecha dd/mm/yyyy a formato Oracle
- `ConvierteTextoComillaAmpersandOracle2(texto)`: Escapa caracteres especiales
- `GEN_QUOTE(texto, char1, char2)`: Reemplaza caracteres en texto
- `INSERTAR_SUPLEMENTO_TSSUPC`: Inserta cláusulas asociadas al suplemento

### Operaciones Relacionadas (Contexto de Uso)

En el flujo real del formulario `FrmFechaPror.frm`, después de ejecutar `INSERTAR_SUPLEMENTO_TCSUPL`, se realizan operaciones adicionales que forman parte de la misma unidad transaccional:

#### Tabla de Operaciones del Flujo Completo

| Orden | Operación | Tabla/Función | Descripción |
|-------|-----------|---------------|-------------|
| 1 | INSERT | `TCSUPL` | Inserción del suplemento (INSERTAR_SUPLEMENTO_TCSUPL) |
| 2 | INSERT | `TSSUPC` | Inserción de cláusulas (INSERTAR_SUPLEMENTO_TSSUPC) |
| 3 | UPDATE | `DTPOLI` | Actualiza fecha de modificación e incrementa número de suplemento |
| 4 | DELETE | `TMPROR` | Limpia registros temporales de prorrateo (PMS_BORRA_TMPROR) |
| 5 | COMMIT | - | Confirma la transacción (SQL_COMMIT) |
| 6 | POST | BDI | Integración con BDI (IntegrarConBDI_ODBC) |

#### UPDATE DTPOLI - Actualización de Póliza

Inmediatamente después de insertar el suplemento, se actualiza la tabla de pólizas:

```vb
' UPDATE DTPOLI - Actualiza fecha de modificación e incrementa número de suplemento
F_STAT = "UPDATE DTPOLI SET "
F_STAT = F_STAT & " POLIFECM = " & Format(Now, "YYYYMMDD")
F_STAT = F_STAT & " ,POLINUSU= " & Trim(Val(G_POLINUSU) + 1)
F_STAT = F_STAT + " WHERE POLICDDE = "
F_STAT = F_STAT + G_POLICDDE + " And POLINPOL = " + G_POLINPOL
HSTMT = SQL_EXEC(G_HDBC, F_STAT, 0)
F_ENDCODE = SQL_END(HSTMT)
```

**SQL Equivalente:**
```sql
UPDATE DTPOLI 
SET POLIFECM = :modifiedAtDate,  -- YYYYMMDD (fecha actual)
    POLINUSU = :supplementNumber  -- G_POLINUSU + 1
WHERE POLICDDE = :delegationCode 
  AND POLINPOL = :policyNumber
```

| Campo | Descripción |
|-------|-------------|
| `POLIFECM` | Fecha de última modificación de la póliza (formato YYYYMMDD) |
| `POLINUSU` | Número de suplemento actual de la póliza (se incrementa en 1) |

## Limpieza de Temporales (PMS_BORRA_TMPROR)

```vb
' Limpia temporales de prorrateo
Call PMS_BORRA_TMPROR(G_POLINPOL, "0")
```

PMS_BORRA_TMPROR lo unic q fa es cridar a una de estes en Database Navigator: Conexión → Schemas → SCOTT → Packages → busca PCK_TARPRO.PMS_BORRA_TMPROR()

Esta llamada elimina los registros temporales de prorrateo de la tabla `TMPROR` para la póliza procesada.

#### Confirmación de Transacción

```vb
' Confirma transacción
ihstmt = SQL_COMMIT(G_HDBC)
```

Todas las operaciones anteriores (INSERT TCSUPL, INSERT TSSUPC, UPDATE DTPOLI, DELETE TMPROR) se confirman en una única transacción.

## Integración con BDI

```vb
' Integración con BDI
Dim sError As String
If Not IntegrarConBDI_ODBC(G_HDBC, cteTipoProcesoPoliza, "635", "", Trim(G_POLINPOL), "", sError) Then
    MsgBox "No se pudo enviar la Integración de la Póliza a BDI", vbCritical, "ATENCION"
End If
```

**Parámetros de IntegrarConBDI_ODBC:**
- `G_HDBC`: Handle de conexión
- `cteTipoProcesoPoliza`: Constante que indica tipo de proceso (póliza)
- `"635"`: Código de delegación (Salud)
- `G_POLINPOL`: Número de póliza

> **Nota:** La integración con BDI se ejecuta DESPUÉS del COMMIT. Si falla, los datos ya están persistidos pero se muestra un mensaje de error al usuario. No hay rollback de la operación principal.

#### ¿Por qué estas operaciones son parte del flujo transaccional?

1. **Consistencia de datos**: El número de suplemento en `DTPOLI.POLINUSU` debe coincidir con el último suplemento insertado en `TCSUPL`
2. **Fecha de modificación**: `DTPOLI.POLIFECM` debe reflejar cuándo se realizó el último cambio en la póliza
3. **Limpieza de temporales**: Los datos de prorrateo en `TMPROR` son específicos de la sesión/operación y deben eliminarse
4. **Atomicidad**: Si alguna operación falla, todas deben revertirse para evitar inconsistencias


### Notas de Migración

1. **Transaccionalidad**: El código legacy no maneja transacciones explícitas. En Spring, envolver en `@Transactional` junto con la llamada a `INSERTAR_SUPLEMENTO_TSSUPC`.

2. **Domicilio normalizado**: La lógica de parsear el domicilio por "#" debe extraerse a un método dedicado o usar un DTO `AddressDto` con campos separados desde el inicio.

3. **Descuentos**: El array `G_Descuentos(0..7)` con 8 grupos debe mapearse a campos individuales o a una estructura embebida.

4. **Variables globales del módulo**: `G_SUPLADAP_SUP`, `G_SUPLINDVTAC_SUP`, etc. deben pasarse como parámetros explícitos en Java, no como estado global.

5. **Formato de fechas**: Convertir de `dd/mm/yyyy` (VB6) a `LocalDate` en Java y usar `DateTimeFormatter` para formatear a `YYYYMMDD` al persistir.

6. **Validación de datos**: Añadir validaciones Jakarta (`@NotNull`, `@Size`) en el DTO de request que no existían en el código legacy.

7. **Entidad existente**: Ya existe `SupplementEntity.java` en el proyecto que mapea la tabla `DTSUPL` (alias de `TCSUPL`). Reutilizar y extender si es necesario.

8. **Cláusulas (TSSUPC)**: Crear entidad y repositorio para `TSSUPC` si no existen, ya que la función delega en `INSERTAR_SUPLEMENTO_TSSUPC` para persistir las cláusulas del suplemento.

9. **⚠️ Flujo transaccional completo**: En Java, todas las operaciones del flujo (INSERT TCSUPL, INSERT TSSUPC, UPDATE DTPOLI, DELETE TMPROR) **deben estar dentro del mismo método `@Transactional`** para garantizar atomicidad:

   ```java
   @Service
   @Transactional
   public class SupplementTransactionalService {
       
       public void createSupplementWithFullFlow(SupplementCreateRequest request, ...) {
           // 1. INSERT TCSUPL + TSSUPC
           supplementCreationService.createSupplement(request, clauseData, ...);
           
           // 2. UPDATE DTPOLI (incrementar número suplemento y fecha modificación)
           policyRepository.updateModificationDateAndSupplementNumber(
               request.getDelegationCode(),
               request.getPolicyNumber(),
               LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE),
               currentSupplementNumber + 1
           );
           
           // 3. DELETE temporales de prorrateo
           prorationTempRepository.deleteByPolicyNumber(request.getPolicyNumber());
           
           // COMMIT implícito al salir del método @Transactional
       }
       
       // Integración BDI se ejecuta DESPUÉS del commit (en otro método o mediante @TransactionalEventListener)
       @TransactionalEventListener(phase = TransactionPhase.AFTER_COMMIT)
       public void handleSupplementCreated(SupplementCreatedEvent event) {
           try {
               bdiIntegrationService.integratePolicyWithBDI(
                   "635", // delegación Salud
                   event.getPolicyNumber()
               );
           } catch (BdiIntegrationException e) {
               log.error("No se pudo enviar la Integración de la Póliza a BDI: {}", e.getMessage());
               // Notificar al usuario pero NO hacer rollback (ya está committed)
           }
       }
   }
   ```

10. **Integración BDI post-commit**: La llamada a `IntegrarConBDI_ODBC` ocurre DESPUÉS del commit. En Spring, usar `@TransactionalEventListener(phase = TransactionPhase.AFTER_COMMIT)` para replicar este comportamiento, o ejecutar la integración en un método separado no transaccional.

---

<!-- Añadir más funciones a continuación -->
