# Referencia de Funciones Legacy VB6

Este documento recopila la documentación de funciones legacy VB6 relevantes para la migración, con su lógica resumida y comportamiento.

---

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

### Firma

```vb
Function GEN_VALFEC(MICONTROL As Control) As Integer
```

### Parámetros

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `MICONTROL` | Control (TextBox) | Control que contiene la fecha a validar |

### Retorno

| Valor | Significado |
|-------|-------------|
| `0` | Fecha válida |
| `-1` | Fecha inválida (muestra mensaje de error) |

### Lógica Resumida

1. **Verifica longitud** = 10 caracteres (`dd/mm/yyyy`)
2. **Verifica separadores** en posiciones 3 y 6 (`/`)
3. **Valida que sean dígitos** en posiciones de día, mes y año
4. **Valida rangos**:
   - Día: 1-31
   - Mes: 1-12
   - Año: 1901-2998
5. **Valida días según el mes**:
   - Febrero: máximo 28 (o 29 si es bisiesto)
   - Meses de 30 días: abril, junio, septiembre, noviembre
   - Meses de 31 días: resto
6. **Si falla**: muestra mensaje "FECHA INCORRECTA" y retorna `-1`

### Ejemplo de Uso

```vb
' Validar fecha antes de procesar
If Trim(TX_Fech.Text) = "" Or GEN_VALFEC(TX_Fech) <> 0 Then
    ' Fecha vacía o inválida
    G_I = PMF_ERROR("GEI024")
    TX_Fech.SetFocus
Else
    ' Continuar con el proceso...
End If
```

### Equivalente Java Propuesto

```java
public boolean isValidDate(String dateStr) {
    if (dateStr == null || dateStr.length() != 10) {
        return false;
    }
    
    try {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        LocalDate date = LocalDate.parse(dateStr, formatter);
        int year = date.getYear();
        return year >= 1901 && year <= 2998;
    } catch (DateTimeParseException e) {
        return false;
    }
}
```

### Notas de Migración

- En Java, usar `LocalDate` con `DateTimeFormatter` maneja automáticamente la validación de días por mes y años bisiestos
- El rango de años (1901-2998) debe mantenerse si es un requisito de negocio
- Considerar si el mensaje de error debe mostrarse en la capa de validación o delegarse al frontend

---

## PMF_ERROR

**Módulo:** `SDGEN1.BAS`

**Propósito:** Muestra mensajes de error personalizados obtenidos de la base de datos y retorna el código de la tecla pulsada por el usuario.

### Firma

```vb
Function PMF_ERROR(CODERROR As String, Optional Incluir As String) As Integer
```

### Parámetros

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `CODERROR` | String | Código del error a buscar (ej: "GEI024") |
| `Incluir` | String (Opcional) | Texto adicional a concatenar al mensaje |


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

### Ejemplo de Uso

```vb
' Validar campo obligatorio
If Trim(TX_Fech.Text) = "" Or GEN_VALFEC(TX_Fech) <> 0 Then
    G_I = PMF_ERROR("GEI024")
    TX_Fech.SetFocus
    Exit Sub
End If

' Confirmación con mensaje adicional
G_I = PMF_ERROR("CLI001", "Póliza: " & numPoliza)
If G_I = 6 Then
    ' Usuario confirmó, guardar cambios
End If
```

### Equivalente Java Propuesto

```java
@Service
public class ErrorMessageService {

    private final ErrorMessageRepository errorRepository;
    
    public ErrorMessageService(ErrorMessageRepository errorRepository) {
        this.errorRepository = errorRepository;
    }
    
    /**
     * Obtiene el mensaje de error desde la base de datos.
     * 
     * @param errorCode Código del error (ej: "GEI024")
     * @return DTO con descripción, tipo y título del mensaje
     */
    public ErrorMessageDto getErrorMessage(String errorCode) {
        return errorRepository.findByCode(errorCode)
            .map(entity -> ErrorMessageDto.builder()
                .code(entity.getCode())
                .description(entity.getDescription())
                .messageType(entity.getMessageType())
                .title(entity.getTitle())
                .build())
            .orElse(ErrorMessageDto.builder()
                .code(errorCode)
                .description("Error desconocido: " + errorCode)
                .messageType(0)
                .title("Error")
                .build());
    }
}

// DTO para la respuesta
@Data
@Builder
public class ErrorMessageDto {
    private String code;
    private String description;
    private Integer messageType;
    private String title;
}

// Entidad JPA
@Entity
@Table(name = "DTMERR")
public class ErrorMessageEntity {
    @Id
    @Column(name = "MERRCDER")
    private String code;
    
    @Column(name = "MERRDSER")
    private String description;
    
    @Column(name = "MERRCDTE")
    private Integer messageType;
    
    @Column(name = "MERRDSCT")
    private String title;
}
```

### Notas de Migración

- **Separación de responsabilidades**: En la arquitectura moderna, la obtención del mensaje (backend) y su visualización (frontend) deben estar separadas
- **API REST**: Exponer endpoint para consultar mensajes de error por código
- **Internacionalización**: Considerar si los mensajes deben soportar múltiples idiomas
- **Tipos de mensaje VB6**: Los valores de `M_TIPOERROR` corresponden a constantes de VB6 para `MsgBox` (ej: `vbOKOnly`, `vbYesNo`, `vbCritical`)
- **Modo lectura**: La lógica de `G_TIPOACC = "R"` debe manejarse en el contexto de permisos del usuario en el sistema nuevo

---

## GRABAR_MOVIMIENTO

**Módulo:** `IBERCAJA.BAS`

**Propósito:** Inserta un registro en la tabla `TCMOVI` para auditoría de movimientos de pólizas. Registra cambios realizados en las pólizas con valores antes/después y metadatos del movimiento.


### SQL Generada Dinámicamente

La función construye la SQL de forma **DINÁMICA**, añadiendo campos opcionales solo si tienen valor:

#### Versión MÍNIMA (solo campos obligatorios)

Cuando `sCliente` está vacío y ningún campo opcional tiene valor:

```sql
INSERT INTO TCMOVI(MOVINUME, MOVITIPO, MOVICDDE, MOVIFECH, MOVINPOL, MOVICDCE, MOVIUSUA, MOVINUSU, MOVICDCL) 
VALUES (SQ_MOVIM.NEXTVAL, 'XX', 635, 'YYYYMMDD', 12345678, 0, 'USUARIO', 1, NULL)
```

#### Versión COMPLETA (todos los campos opcionales informados)

Cuando todos los campos opcionales tienen valor:

```sql
INSERT INTO TCMOVI(MOVINUME, MOVITIPO, MOVICDDE, MOVIFECH, MOVINPOL, MOVICDCE, MOVIUSUA, MOVINUSU, MOVICDCL, MOVIANTE, MOVIDESP, MOVIFECH_EFECTO, MOVITAREA_SGO) 
VALUES (SQ_MOVIM.NEXTVAL, 'XX', 635, 'YYYYMMDD', 12345678, 0, 'USUARIO', 1, 12345, 'Valor anterior', 'Valor nuevo', 'YYYYMMDD', 'EXP123456')
```


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
| `sTipo` | String | Código del tipo de movimiento (ver constantes G_MOV_*) |
| `sDelegacion` | String | Código de delegación (635=Salud, 634=Decesos) |
| `sFecha` | String | Fecha del movimiento en formato YYYYMMDD |
| `sPoliza` | String | Número de póliza |
| `sCertificado` | String | Número de certificado |
| `sUsuario` | String | Usuario que realiza el movimiento |
| `sSuplemento` | String | Número de suplemento |
| `sCliente` | String | Código de cliente (puede ser NULL)
| `sAntes` | String | Valor antes del cambio |
| `sDespues` | String | Valor después del cambio |
| `sFecha_Efecto` | String (Opcional) | Fecha de efecto del movimiento en formato YYYYMMDD |
| `sNumExpSGO` | String (Opcional) | Número de expediente SGO |

### Retorno

| Valor | Significado |
|-------|-------------|
| `0` | Movimiento grabado correctamente |
| `-1` | Error al grabar el movimiento |

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

### Estructura de Base de Datos

**Tabla:** `TCMOVI`

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `MOVITIPO` | VARCHAR2(2) | Tipo de movimiento (G_MOV_*) |
| `MOVIDELE` | VARCHAR2(3) | Código de delegación |
| `MOVIFECH` | VARCHAR2(8) | Fecha del movimiento (YYYYMMDD) |
| `MOVIPOLI` | VARCHAR2(10) | Número de póliza |
| `MOVICERT` | VARCHAR2(5) | Número de certificado |
| `MOVIUSUA` | VARCHAR2(10) | Usuario que realiza el movimiento |
| `MOVISUPL` | VARCHAR2(3) | Número de suplemento |
| `MOVICLIE` | VARCHAR2(10) | Código de cliente |
| `MOVIANTE` | VARCHAR2(200) | Valor antes del cambio |
| `MOVIDESP` | VARCHAR2(200) | Valor después del cambio |
| `MOVIFEFE` | VARCHAR2(8) | Fecha de efecto (YYYYMMDD) |
| `MOVIEXPS` | VARCHAR2(15) | Número de expediente SGO |
| `MOVIFGRA` | DATE | Fecha/hora de grabación (SYSDATE) |

### Código VB6 Original

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
    
    On Error GoTo ErrorHandler
    
    Dim sSQL As String
    Dim sClienteSQL As String
    Dim sFechaEfectoSQL As String
    Dim sNumExpSQL As String
    
    ' Tratar valores NULL
    If Trim(sCliente) = "" Then
        sClienteSQL = "NULL"
    Else
        sClienteSQL = "'" & sCliente & "'"
    End If
    
    If IsMissing(sFecha_Efecto) Or Trim(sFecha_Efecto) = "" Then
        sFechaEfectoSQL = "NULL"
    Else
        sFechaEfectoSQL = "'" & sFecha_Efecto & "'"
    End If
    
    If IsMissing(sNumExpSGO) Or Trim(sNumExpSGO) = "" Then
        sNumExpSQL = "NULL"
    Else
        sNumExpSQL = "'" & sNumExpSGO & "'"
    End If
    
    ' Truncar valores si exceden longitud máxima
    sAntes = Left(Trim(sAntes), 200)
    sDespues = Left(Trim(sDespues), 200)
    
    sSQL = "INSERT INTO TCMOVI " & _
           "(MOVITIPO, MOVIDELE, MOVIFECH, MOVIPOLI, MOVICERT, " & _
           "MOVIUSUA, MOVISUPL, MOVICLIE, MOVIANTE, MOVIDESP, " & _
           "MOVIFEFE, MOVIEXPS, MOVIFGRA) " & _
           "VALUES (" & _
           "'" & sTipo & "', " & _
           "'" & sDelegacion & "', " & _
           "'" & sFecha & "', " & _
           "'" & sPoliza & "', " & _
           "'" & sCertificado & "', " & _
           "'" & sUsuario & "', " & _
           "'" & sSuplemento & "', " & _
           sClienteSQL & ", " & _
           "'" & Replace(sAntes, "'", "''") & "', " & _
           "'" & Replace(sDespues, "'", "''") & "', " & _
           sFechaEfectoSQL & ", " & _
           sNumExpSQL & ", " & _
           "SYSDATE)"
    
    GEN_INSUPDDEL sSQL
    
    GRABAR_MOVIMIENTO = 0
    Exit Function
    
ErrorHandler:
    ' Registrar error pero no interrumpir el flujo principal
    Call RegistrarErrorLog("GRABAR_MOVIMIENTO", Err.Description)
    GRABAR_MOVIMIENTO = -1
End Function
```

### Lógica Resumida

1. **Manejo de valores opcionales**: Convierte strings vacíos a `NULL` para cliente, fecha efecto y número expediente
2. **Truncamiento de valores**: Limita `sAntes` y `sDespues` a 200 caracteres
3. **Escape de caracteres**: Reemplaza comillas simples por dobles en los valores antes/después
4. **Inserción**: Ejecuta INSERT en `TCMOVI` con `SYSDATE` para la fecha de grabación
5. **Manejo de errores**: Captura errores sin interrumpir el flujo principal de la aplicación

### Ejemplos de Uso Reales del Código Legacy

```vb
' Ejemplo 1: Grabar movimiento de cambio de agente
Call GRABAR_MOVIMIENTO(G_MOV_CAMBIO_AGENTE, _
                       "635", _
                       Format(Now, "YYYYMMDD"), _
                       G_Poliza, _
                       G_Certificado, _
                       G_Usuario, _
                       G_Suplemento, _
                       G_Cliente, _
                       "Agente anterior: " & sAgenteAnterior, _
                       "Agente nuevo: " & sAgenteNuevo)

' Ejemplo 2: Grabar movimiento de alta de asegurado con fecha de efecto
Call GRABAR_MOVIMIENTO(G_MOV_ALTA_ASEGURADO, _
                       "635", _
                       Format(Now, "YYYYMMDD"), _
                       G_Poliza, _
                       G_Certificado, _
                       G_Usuario, _
                       G_Suplemento, _
                       G_Cliente, _
                       "", _
                       "NIF: " & sNifAsegurado & " - Nombre: " & sNombreAsegurado, _
                       sFechaEfecto)

' Ejemplo 3: Grabar movimiento de modificación con expediente SGO
Call GRABAR_MOVIMIENTO(G_MOV_MODIFICACION_POLIZA, _
                       "635", _
                       Format(Now, "YYYYMMDD"), _
                       G_Poliza, _
                       G_Certificado, _
                       G_Usuario, _
                       G_Suplemento, _
                       "", _
                       "Dirección anterior: " & sDireccionAnterior, _
                       "Dirección nueva: " & sDireccionNueva, _
                       sFechaEfecto, _
                       sNumExpedienteSGO)

' Ejemplo 4: Grabar movimiento de cambio de forma de pago
Call GRABAR_MOVIMIENTO(G_MOV_CAMBIO_FORMA_PAGO, _
                       G_Delegacion, _
                       Format(Now, "YYYYMMDD"), _
                       txtPoliza.Text, _
                       txtCertificado.Text, _
                       G_UsuarioActual, _
                       "000", _
                       txtCliente.Text, _
                       "Forma pago: " & sFormaPagoAnterior, _
                       "Forma pago: " & sFormaPagoNueva)
```

### Equivalente Java Propuesto

```java
// Entidad JPA
@Entity
@Table(name = "TCMOVI")
public class MovementAuditEntity {
    
    @Id
    @GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "movi_seq")
    @SequenceGenerator(name = "movi_seq", sequenceName = "TCMOVI_SEQ", allocationSize = 1)
    private Long id;
    
    @Column(name = "MOVITIPO", length = 2, nullable = false)
    private String movementType;
    
    @Column(name = "MOVIDELE", length = 3, nullable = false)
    private String delegationCode;
    
    @Column(name = "MOVIFECH", length = 8, nullable = false)
    private String movementDate;
    
    @Column(name = "MOVIPOLI", length = 10, nullable = false)
    private String policyNumber;
    
    @Column(name = "MOVICERT", length = 5, nullable = false)
    private String certificateNumber;
    
    @Column(name = "MOVIUSUA", length = 10, nullable = false)
    private String userId;
    
    @Column(name = "MOVISUPL", length = 3, nullable = false)
    private String supplementNumber;
    
    @Column(name = "MOVICLIE", length = 10)
    private String clientCode;
    
    @Column(name = "MOVIANTE", length = 200)
    private String valueBefore;
    
    @Column(name = "MOVIDESP", length = 200)
    private String valueAfter;
    
    @Column(name = "MOVIFEFE", length = 8)
    private String effectiveDate;
    
    @Column(name = "MOVIEXPS", length = 15)
    private String sgoExpedientNumber;
    
    @Column(name = "MOVIFGRA", nullable = false)
    @CreationTimestamp
    private LocalDateTime recordDate;
}

// Enum para tipos de movimiento
public enum MovementType {
    ALTA_POLIZA("01", "Alta de póliza"),
    BAJA_POLIZA("02", "Baja de póliza"),
    MODIFICACION_POLIZA("03", "Modificación de póliza"),
    SUPLEMENTO("04", "Suplemento"),
    ANULACION("05", "Anulación"),
    REHABILITACION("06", "Rehabilitación"),
    CAMBIO_TOMADOR("07", "Cambio de tomador"),
    CAMBIO_AGENTE("08", "Cambio de agente"),
    CAMBIO_DOMICILIO("09", "Cambio de domicilio"),
    CAMBIO_FORMA_PAGO("10", "Cambio de forma de pago"),
    CAMBIO_CUENTA("11", "Cambio de cuenta"),
    ALTA_ASEGURADO("12", "Alta de asegurado"),
    BAJA_ASEGURADO("13", "Baja de asegurado"),
    MODIFICACION_ASEGURADO("14", "Modificación de asegurado"),
    CAMBIO_COBERTURA("15", "Cambio de cobertura"),
    CAMBIO_TARIFA("16", "Cambio de tarifa"),
    CAMBIO_DESCUENTO("17", "Cambio de descuento"),
    RENOVACION("18", "Renovación"),
    EMISION_RECIBO("19", "Emisión de recibo"),
    ANULACION_RECIBO("20", "Anulación de recibo");
    
    private final String code;
    private final String description;
    
    MovementType(String code, String description) {
        this.code = code;
        this.description = description;
    }
    
    public String getCode() {
        return code;
    }
    
    public String getDescription() {
        return description;
    }
    
    public static MovementType fromCode(String code) {
        for (MovementType type : values()) {
            if (type.code.equals(code)) {
                return type;
            }
        }
        throw new IllegalArgumentException("Unknown movement type code: " + code);
    }
}

// DTO de request
@Data
@Builder
public class MovementAuditRequest {
    @NotNull
    private MovementType movementType;
    
    @NotBlank
    @Size(max = 3)
    private String delegationCode;
    
    @NotBlank
    @Pattern(regexp = "\\d{8}", message = "Date must be in YYYYMMDD format")
    private String movementDate;
    
    @NotBlank
    @Size(max = 10)
    private String policyNumber;
    
    @NotBlank
    @Size(max = 5)
    private String certificateNumber;
    
    @NotBlank
    @Size(max = 10)
    private String userId;
    
    @NotBlank
    @Size(max = 3)
    private String supplementNumber;
    
    @Size(max = 10)
    private String clientCode;
    
    @Size(max = 200)
    private String valueBefore;
    
    @Size(max = 200)
    private String valueAfter;
    
    @Pattern(regexp = "\\d{8}", message = "Effective date must be in YYYYMMDD format")
    private String effectiveDate;
    
    @Size(max = 15)
    private String sgoExpedientNumber;
}

// Repository
@Repository
public interface MovementAuditRepository extends JpaRepository<MovementAuditEntity, Long> {
    
    List<MovementAuditEntity> findByPolicyNumberAndCertificateNumberOrderByRecordDateDesc(
        String policyNumber, String certificateNumber);
    
    List<MovementAuditEntity> findByPolicyNumberAndMovementTypeOrderByRecordDateDesc(
        String policyNumber, String movementType);
}

// Service
@Service
@Slf4j
public class MovementAuditService {
    
    private final MovementAuditRepository repository;
    
    public MovementAuditService(MovementAuditRepository repository) {
        this.repository = repository;
    }
    
    /**
     * Records a policy movement for audit purposes.
     * Equivalent to VB6 GRABAR_MOVIMIENTO function.
     *
     * @param request Movement audit data
     * @return true if recorded successfully, false otherwise
     */
    @Transactional
    public boolean recordMovement(MovementAuditRequest request) {
        try {
            MovementAuditEntity entity = MovementAuditEntity.builder()
                .movementType(request.getMovementType().getCode())
                .delegationCode(request.getDelegationCode())
                .movementDate(request.getMovementDate())
                .policyNumber(request.getPolicyNumber())
                .certificateNumber(request.getCertificateNumber())
                .userId(request.getUserId())
                .supplementNumber(request.getSupplementNumber())
                .clientCode(StringUtils.hasText(request.getClientCode()) 
                    ? request.getClientCode() : null)
                .valueBefore(truncateToMaxLength(request.getValueBefore(), 200))
                .valueAfter(truncateToMaxLength(request.getValueAfter(), 200))
                .effectiveDate(StringUtils.hasText(request.getEffectiveDate()) 
                    ? request.getEffectiveDate() : null)
                .sgoExpedientNumber(StringUtils.hasText(request.getSgoExpedientNumber()) 
                    ? request.getSgoExpedientNumber() : null)
                .build();
            
            repository.save(entity);
            
            log.info("Movement recorded: type={}, policy={}, certificate={}, user={}",
                request.getMovementType().getCode(),
                request.getPolicyNumber(),
                request.getCertificateNumber(),
                request.getUserId());
            
            return true;
            
        } catch (Exception e) {
            log.error("Error recording movement for policy {}: {}", 
                request.getPolicyNumber(), e.getMessage(), e);
            return false;
        }
    }
    
    private String truncateToMaxLength(String value, int maxLength) {
        if (value == null) {
            return null;
        }
        String trimmed = value.trim();
        return trimmed.length() > maxLength 
            ? trimmed.substring(0, maxLength) 
            : trimmed;
    }
}
```

### Notas de Migración

1. **Formato de fechas**: El código legacy usa formato `YYYYMMDD` como String. Considerar migrar a `LocalDate` en el modelo interno y convertir solo para persistencia si la tabla no se puede modificar.

2. **Truncamiento de valores**: Mantener la lógica de truncar `valueBefore` y `valueAfter` a 200 caracteres para compatibilidad con la estructura de tabla existente.

3. **Valores NULL**: El código VB6 trata strings vacíos como NULL. En Java, usar `StringUtils.hasText()` para la misma lógica.

4. **Manejo de errores silencioso**: El código legacy captura errores sin propagarlos para no interrumpir el flujo principal. En la migración, mantener este comportamiento pero asegurar logging adecuado.

5. **Delegaciones**: Los códigos de delegación (635=Salud, 634=Decesos) podrían externalizarse a configuración o enumeración.

6. **Transaccionalidad**: El movimiento de auditoría debería grabarse en la misma transacción que la operación principal para garantizar consistencia. Considerar usar `@Transactional(propagation = Propagation.MANDATORY)` si debe ejecutarse dentro de una transacción existente.

7. **Seguridad**: El campo `sUsuario` viene del contexto de sesión VB6. En Spring, obtener del `SecurityContext` o inyectar mediante AOP.

8. **Código existente en el proyecto**: Verificar si `MovementEntity` y `MovementService` existentes en el proyecto pueden extenderse o si se requiere una nueva entidad específica para auditoría.

---

## INSERTAR_SUPLEMENTO_TCSUPL

**Módulo:** `mdlSuplementos.bas`

**Propósito:** Inserta un registro de suplemento en la tabla TCSUPL con todos los datos históricos de una póliza/certificado. Construye dinámicamente el INSERT según si hay datos de domicilio nuevos o antiguos. Delega en INSERTAR_SUPLEMENTO_TSSUPC para las cláusulas asociadas.

### Firma

```vb
Public Sub INSERTAR_SUPLEMENTO_TCSUPL(ByVal G_POCE2CDDE As String, ByVal G_POCE2NPOL As String, _
                           ByVal G_POCE2CDCE As String, ByVal G_POCE2NUSU As String, ByVal G_CER As String, ByVal G_POLI2ESPA As String, ByVal G_POCE2ALBA As String, _
                           ByVal G_POCE2FECB As String, ByVal G_POCE2FECA As String, ByVal G_POCE2FEBA As String, ByVal G_POLI2CDPT As String, _
                           ByVal G_POCE2CDTA As String, ByVal G_POCE2FOPA As String, ByVal G_POCE2TIPA As String, ByVal G_POCE2NUPE As String, _
                           ByVal G_POLI2IDCP As String, ByVal G_POLI2IDCO As String, ByVal G_POLI2CDTR As String, ByVal G_POLI2IDEX As String, _
                           ByVal G_POCE2COBR As String, ByVal G_POCE2AGTA As String, ByVal G_POCE2AGTB As String, _
                           ByVal G_CAT_TIPO_ANT As String, ByVal G_POCE2PRNT As String, ByVal G_POCE2PRNE As String, ByVal G_POCE2RECA As String, _
                           ByVal G_POCE2RECE As String, ByVal G_POCE2IMPT As String, ByVal G_POCE2IMPE As String, ByVal G_POCE2TORE As String, _
                           ByVal G_POCE2MOBA As String, ByVal G_POCE2DOMI As String, ByVal G_POCE2CDPS As String, ByVal G_POCE2CDPO As String, _
                           ByVal G_POCE2TFNO As String, ByVal G_POCE2CRNT As String, ByVal G_POCE2CRNE As String, ByVal G_POCE2CECA As String, _
                           ByVal G_POCE2CECE As String, ByVal G_POCE2CMPT As String, ByVal G_POCE2CMPE As String, ByVal G_POCE2CORE As String, _
                           ByVal G_POCE2IPUN As String, ByVal G_POCE2IPUE As String, ByVal G_POCE2CPUN As String, ByVal G_POCE2CPUE As String, _
                           ByVal G_Fecha_Asegurado As String, ByVal G_POLI2IDMA As String, ByVal G_POLI2NUCE As String, ByVal G_POLI2CDRP As String, _
                           ByRef G_POCEFECM As String, ByRef strSuplemento_DTSUAS As String, _
                           ByVal G_POCEPROVINCIA_TARIFICACION As String, ByVal G_POCESWPROVINCIA As String, ByVal G_POCESWPRODUCCION As String, _
                           ByVal G_POCESWTARIFA As String, ByVal G_POCESWDCTO As String, ByVal G_POCEDCTONUMPERSONAS As String, ByVal G_POCERECFORMAPAGO As String, _
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

### Retorno

No retorna valor (Sub). Los valores de salida se devuelven por referencia en `G_POCEFECM` y `strSuplemento_DTSUAS`.

### Estructura de Base de Datos

**Tabla principal:** `TCSUPL`

Campos principales insertados (70+ columnas):

| Campo | Descripción |
|-------|-------------|
| `SUPLCDDE` | Código de delegación |
| `SUPLNPOL` | Número de póliza |
| `SUPLCDCE` | Código de certificado |
| `SUPLNUSU` | Número de suplemento |
| `SUPLTIPO` | Tipo de entidad (01/02/03) |
| `SUPLSITP` | Estado previo póliza |
| `SUPLSITC` | Estado actual |
| `SUPLFECA` | Fecha creación suplemento (YYYYMMDD) |
| `SUPLFECB` | Fecha efecto B |
| `SUPLFECC` | Fecha efecto C |
| `SUPLFEBA` | Fecha baja |
| `SUPLCDPT` | Código producto |
| `SUPLCDTA` | Código tarifa |
| `SUPLFOPA` | Forma de pago |
| `SUPLPRNT/SUPLPRNE` | Primas netas (total/extranjero) |
| `SUPLIMPT/SUPLIMPE` | Impuestos (total/extranjero) |
| `SUPLTORE` | Total recibo |
| `SUPLDESC_G01..G08` | 8 grupos de descuento |
| `SUP_CDG_TIPOVIA..SUP_CVIA_INE` | Campos de domicilio normalizado |
| `SUPLINDVTAC` | Indicador venta cruzada |
| `SUPLINDPRIC` | Indicador pricing |

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

### Ejemplo de Uso Real del Código Legacy

```vb
' Desde Generar_Suplemento() después de cargar datos de la póliza/certificado:
mdlSuplementos.INSERTAR_SUPLEMENTO_TCSUPL sCDDE, sNpol, _
                                            sCDCE, sNusu, _
                                            sTipo, sESPA, sALBA, _
                                            sFECB, sFECA, _
                                            sFEBA, sCDPT, _
                                            sCDTA, sFOPA, _
                                            sTIPA, sNUPE, _
                                            sIDCP, sIDCO, _
                                            sCDTR, sIDEX, sCOBR, _
                                            sAGTA, sAGTB, _
                                            sINSP, sPRNT, _
                                            sPRNE, sRECA, _
                                            sRECE, sIMPT, _
                                            sIMPE, sTORE, sMOBA, sDOMI, _
                                            sCDPS, sCDPO, sTFNO, sCRNT, _
                                            sCRNE, sCECA, sCECE, sCMPT, _
                                            sCMPE, sCORE, sIPUN, sIPUE, _
                                            sCPUN, sCPUE, _
                                            Format(Now, "dd/mm/yyyy"), sIDMA, sNUCE, sCDRP, _
                                            sFECM, sSUAS, sPROVINCIA_TARIFICACION, sSWPROVINCIA, _
                                            sSWPRODUCCION, sSWTARIFA, sSWDCTO, sDCTONUMPERSONAS, _
                                            sRECFORMAPAGO, aSubDescuentos
```

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
SET POLIFECM = :fechaModificacion,  -- YYYYMMDD (fecha actual)
    POLINUSU = :nuevoNumeroSuplemento  -- G_POLINUSU + 1
WHERE POLICDDE = :codigoDelegacion 
  AND POLINPOL = :numeroPoliza
```

| Campo | Descripción |
|-------|-------------|
| `POLIFECM` | Fecha de última modificación de la póliza (formato YYYYMMDD) |
| `POLINUSU` | Número de suplemento actual de la póliza (se incrementa en 1) |

#### Limpieza de Temporales (PMS_BORRA_TMPROR)

```vb
' Limpia temporales de prorrateo
Call PMS_BORRA_TMPROR(G_POLINPOL, "0")
```

Esta llamada elimina los registros temporales de prorrateo de la tabla `TMPROR` para la póliza procesada.

#### Confirmación de Transacción

```vb
' Confirma transacción
ihstmt = SQL_COMMIT(G_HDBC)
```

Todas las operaciones anteriores (INSERT TCSUPL, INSERT TSSUPC, UPDATE DTPOLI, DELETE TMPROR) se confirman en una única transacción.

#### Integración con BDI

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

### Equivalente Java Propuesto

```java
// DTO de request para creación de suplemento
@Data
@Builder
public class SupplementCreateRequest {
    @NotBlank
    private String delegationCode;
    
    @NotBlank
    private String policyNumber;
    
    @NotBlank
    private String certificateCode;
    
    @NotBlank
    private String supplementNumber;
    
    @NotBlank
    @Pattern(regexp = "01|02|03", message = "Entity type must be 01, 02, or 03")
    private String entityType;
    
    private String previousStatus;
    private String currentStatus;
    
    private LocalDate effectDateB;
    private LocalDate effectDateA;
    private LocalDate terminationDate;
    
    private String productCode;
    private String tariffCode;
    private String paymentMethod;
    private String paymentType;
    private Integer insuredCount;
    
    // Premium amounts
    private BigDecimal netPremiumTotal;
    private BigDecimal netPremiumForeign;
    private BigDecimal surchargeTotal;
    private BigDecimal surchargeForeign;
    private BigDecimal taxTotal;
    private BigDecimal taxForeign;
    private BigDecimal totalAnnualReceipt;
    
    // Address - can be normalized or legacy format
    private AddressDto address;
    
    // Discount groups (G01-G08)
    private List<BigDecimal> discountGroups;
    
    // Indicators
    private String crossSellIndicator;
    private String pricingIndicator;
    
    // Rating province
    private String ratingProvinceCode;
}

// DTO para dirección normalizada
@Data
@Builder
public class AddressDto {
    private String streetTypeCode;
    private String streetName;
    private String streetNumber;
    private String portal;
    private String block;
    private String staircase;
    private String floor;
    private String door;
    private String additionalInfo;
    private String ineCityCode;
    private String ineStreetCode;
    private String postalCode;
    private String populationCode;
}

// Service para creación de suplementos
@Service
@Slf4j
@Transactional
public class SupplementCreationService {
    
    private final SupplementRepository supplementRepository;
    private final SupplementClauseRepository clauseRepository;
    private final SupplementMapper supplementMapper;
    
    public SupplementCreationService(SupplementRepository supplementRepository,
                                     SupplementClauseRepository clauseRepository,
                                     SupplementMapper supplementMapper) {
        this.supplementRepository = supplementRepository;
        this.clauseRepository = clauseRepository;
        this.supplementMapper = supplementMapper;
    }
    
    /**
     * Creates a supplement record in TCSUPL with all historical data.
     * Equivalent to VB6 INSERTAR_SUPLEMENTO_TCSUPL function.
     *
     * @param request Supplement creation data
     * @param clauseData Pre-loaded clause data (equivalent to MAT_TSPOPC)
     * @param adaptationValue Value from G_SUPLADAP_SUP
     * @param territorialityValue Value from G_SUPLTERRITORIALIDAD_SUP
     * @return Created supplement number
     */
    public String createSupplement(SupplementCreateRequest request,
                                   List<ClauseData> clauseData,
                                   String adaptationValue,
                                   String territorialityValue) {
        
        log.info("Creating supplement for policy {}, certificate {}",
            request.getPolicyNumber(), request.getCertificateCode());
        
        // Build entity from request
        SupplementEntity entity = supplementMapper.toEntity(request);
        
        // Set creation date (equivalent to Format(Now, "YYYYMMDD"))
        entity.setCreationDate(LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE));
        
        // Set adaptation and territoriality values
        entity.setAdaptationValue(adaptationValue);
        entity.setTerritorialityValue(territorialityValue);
        
        // Set discount groups (G01-G08)
        if (request.getDiscountGroups() != null && request.getDiscountGroups().size() == 8) {
            entity.setDiscountG01(request.getDiscountGroups().get(0));
            entity.setDiscountG02(request.getDiscountGroups().get(1));
            entity.setDiscountG03(request.getDiscountGroups().get(2));
            entity.setDiscountG04(request.getDiscountGroups().get(3));
            entity.setDiscountG05(request.getDiscountGroups().get(4));
            entity.setDiscountG06(request.getDiscountGroups().get(5));
            entity.setDiscountG07(request.getDiscountGroups().get(6));
            entity.setDiscountG08(request.getDiscountGroups().get(7));
        }
        
        // Save main supplement record
        supplementRepository.save(entity);
        
        // Insert clauses if NOT collective policy (G_CER <> "02")
        if (!"02".equals(request.getEntityType())) {
            insertSupplementClauses(
                request.getPolicyNumber(),
                request.getCertificateCode(),
                request.getSupplementNumber(),
                clauseData
            );
        }
        
        log.info("Supplement {} created successfully", request.getSupplementNumber());
        
        return request.getSupplementNumber();
    }
    
    /**
     * Inserts clause records in TSSUPC.
     * Equivalent to VB6 INSERTAR_SUPLEMENTO_TSSUPC function.
     */
    private void insertSupplementClauses(String policyNumber,
                                         String certificateCode,
                                         String supplementNumber,
                                         List<ClauseData> clauseData) {
        for (ClauseData clause : clauseData) {
            if (clause.getClauseCode() != null && !clause.getClauseCode().isEmpty()) {
                SupplementClauseEntity clauseEntity = SupplementClauseEntity.builder()
                    .policyNumber(policyNumber)
                    .certificateCode(certificateCode)
                    .supplementNumber(supplementNumber)
                    .clauseCode(clause.getClauseCode())
                    .relationCode(clause.getRelationCode())
                    .coverageType(clause.getCoverageType())
                    .premiumAmount(clause.getPremiumAmount())
                    .startDate(clause.getStartDate())
                    .endDate(clause.getEndDate())
                    .previousAmount(clause.getPreviousAmount())
                    .previousDate(clause.getPreviousDate())
                    .receiptWithoutDiscount(clause.getReceiptWithoutDiscount())
                    .premiumWithoutDiscount(clause.getPremiumWithoutDiscount())
                    .tariffDate(clause.getTariffDate())
                    .build();
                
                clauseRepository.save(clauseEntity);
            }
        }
    }
}

// Repository usando la entidad existente SupplementEntity
@Repository
public interface SupplementRepository extends JpaRepository<SupplementEntity, SupplementPK> {
    
    List<SupplementEntity> findByPolicyNumberOrderBySupplementNumberDesc(Integer policyNumber);
    
    Optional<SupplementEntity> findFirstByPolicyNumberAndCertificateCodeOrderBySupplementNumberDesc(
        Integer policyNumber, Integer certificateCode);
}
```

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
