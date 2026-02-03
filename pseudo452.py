from datetime import date

def tarificar_click(date_input_value: date):
    if date_input_value is None:
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

    if swProrrateos is "N" or None: #no se de donde sale swProrrateos
        emitirProrrateo = False

    if codCliente == tomadorPoliza: bCambioTomador = True


    codig

    pass

def grabar_movimiento():
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
