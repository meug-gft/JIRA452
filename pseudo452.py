from datetime import date

def tarificar_click(date_input_value: date):

    if not is_valid_date(date_input_value): #is_valid_date() controla validez de fecha y que no esté vacío. esta función es GEN_VALFEC() -> GENRT.bas:1005
        pmf_error("GEI024")
    else:
        codEspecialidad: int 
        nPol: int
        fechaAsegurado: date
        G_SW_RETARIFICAR: bool
        G_POLINUSU: int

        if G_TECLA_PRORRATEO == 6:

            pms_tarttoind(muchisimos_parametros = None)

            grabar_movimiento()

            insertar_suplemento_tcsupl(muchos_parametros = None)

            sql_query = '''
                UPDATE DTPOLI SET POLIFECM = CURRENT_DATE(), POLINUSU = (:G_POLINUSU)+1 
                WHERE POLICDDE = :delegationCode
                AND POLINPOL =  :nPol
            '''

            pms_borra_tmpror(nPol, "0")

            #bdi_integration()

        elif G_TECLA_PRORRATEO == 7:
            ROLLBACK()
        
        elif G_TECLA_PRORRATEO is not 2:
            G_SW_CAMBIO = 0
                    
            # Unload FrmFechaPror
            G_SW_CAMBIO = 0
            G_SW_SUPLEMENTO = 0
            G_Sw_Salir = 0
            G_SW_SALVADO_P = 0
            # Unload DCAS015
            # Unload DCAS017
                    
            G_TECLA_PRORRATEO = 0
            G_SW_RETARIFICAR = False
            # DCAS001.Show
    return

def pms_tarttoind(muchisimos_parametros: object):

    user: object = ENV("USR")
    emitirProrrateo :bool = True
    swProrrateos: str
    bCambioTomador: bool
    p_codRamo: str
    proceedTarification: bool

    if swProrrateos is "N" or None: #no se de donde sale swProrrateos
        emitirProrrateo = False

    if codCliente == tomadorPoliza: bCambioTomador = True

    nifTomador = muchisimos_parametros.nif #este campo viene de DCAS015, deberá venir como parámetro de entrada en el endpoint

    formaPago = muchisimos_parametros.CB_Fopa[70:72] #este campo viene de DCAS015, deberá venir como parámetro de entrada en el endpoint

    # Comprobación de modificación de los campos que afectan a la tarificación
    # TODAS ESTAS VARIABLES SON GLOBALES
    if G_Sw_Abrir == -1 and G_Sw_Cob_Esp == -1:
        G_SW_TARMTTO == 1
    elif G_Sw_Abrir == -1 and G_SWSUPL == -1:
        G_SW_TARMTTO == 1
    elif G_POLI2IDEX is not G_CHECK:
        G_SW_TARMTTO == 1
    elif G_POLI2CDTA is not muchisimos_parametros.TX_CDTA.Text: #este campo viene de DCAS015, deberá venir como parámetro de entrada en el endpoint
        G_SW_TARMTTO == 1
    elif G_POLI2FVTAR is not G_POLIFVTAR:
        G_SW_TARMTTO == 1
    elif G_POLI2FOPA is not Forma_Pago:
        G_SW_TARMTTO == 1
    elif G_POLI2TIPA is not muchisimos_parametros.CB_Tipa[71::2]: #este campo viene de DCAS015, deberá venir como parámetro de entrada en el endpoint
        G_SW_TARMTTO == 1
    elif G_POLI2FECM is not G_POLIFECM:
        G_SW_TARMTTO == 1
    elif G_POLI2FECB is not muchisimos_parametros.TX_Fech(2).Text: 
        G_SW_TARMTTO == 1
    #p207 Un cambio en la provincia de tarificaci�n produce una retarificaci�n.
    elif G_POLI2PROVINCIA_TARIFICACION is not muchisimos_parametros.Cmb_ProvinciaTar[71::2]: #este campo viene de DCAS017, deberá venir como parámetro de entrada en el endpoint
        G_SW_TARMTTO = 1
    elif cambios_descuento(G_GruposDescuento, G_Grupos2Descuento):
        G_SW_TARMTTO = 1
    elif G_Grupos2Descuento(4).Valor is not G_GruposDescuento(4).Valor:
        #El descuento de retenci�n ha podido cambiar manualmente
        G_SW_TARMTTO = 1
    #P614.50 Control de cambio de valor de descuento comercial
    elif G_Grupos2Descuento(3).Valor is not G_GruposDescuento(3).Valor:
        #El descuento de comercial ha podido cambiar manualmente
        G_SW_TARMTTO = 1
    elif G_POLI2ESPA is not muchisimos_parametros.CB_Espa.Text[71::2]:
        G_SW_TARMTTO = 1
    
    if G_SW_RETARIFICAR: G_SW_TARMTTO = 1

    #Recuperaci�n de los datos antiguos de DTPOLI
    sql_query = '''
        SELECT 
            POLICDDE, POLINPOL, POLIFECA, POLICDPT,
            POLIFOPA, POLIIDCO, POLICOBR, POLIPRNT, 
            POLIPRNE, POLIRECA, POLIRECE, POLIIMPT, 
            POLIIMPE, POLITORE, POLIIPUN, POLIIPUE, 
            POLINUPE, POLIFEER, null POLITIPO_DCTO, 
            POLITREN, POLIFEVE
        FROM DTPOLI
        WHERE 
            POLICDDE = :delegationCode 
            AND POLINPOL = :policyNumber
        FOR UPDATE'''


    # si hay algun cambio entra a retarificar
    if proceedTarification or (not proceedTarification and bCambioTomador):
        G_POLI2FOPA:str ## no se lo que es G_POLI2FOPA
        bCambioFormaPago: bool
        
        #Comprueba el suplemento realizado
        if G_POLI2FOPA is not formaPago: bCambioFormaPago = True




    pass

def grabar_movimiento():
    pass

def cambios_descuento(G_GruposDescuento:list, G_Grupos2Descuento:list): 
    pass

def insertar_suplemento_tcsupl(muchos_parametros: object):
    pass

def pms_borra_tmpror(policyNumber: int, certificateNumber: str):
    pass

def pmf_error(errorCode: str):
    sql_query = ''' 
    SELECT 
        MERRCDER as errorCode, 
        MERRDSER as errorDescription, 
        MERRCDTE as errorTypeCode, 
        MERRDSCT as shortDescription 
    FROM DTMERR 
    WHERE
        MERRCDER = :errorCode'''
