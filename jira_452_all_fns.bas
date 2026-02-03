''
'' FrmFechaPror.frm()
''
Sub Cmd_Tarificar_Click()
    Dim IndProrratear As Boolean
    Dim sql As String
    Dim ihstmt As Long
    Dim Fetch As Long
    Dim ENDCODE As Integer
    Dim F_STAT As String
    Dim A1 As String
    Dim b As String
    Dim sFECM As String
    Dim strSuplemento_DTSUAS As String

    If Trim(TX_Fech.Text) = "" Or GEN_VALFEC(TX_Fech) <> 0 Then
            Screen.MousePointer = 0
            G_I = PMF_ERROR("GEI024")
            TX_Fech.SetFocus
    Else

       A1 = G_POLICDDE
       b = G_POLINPOL
       G_Fecha_Asegurado = TX_Fech.Text
       G_SW_RETARIFICAR = True
       PMS_TARMTTOIND
       If G_TECLA_PRORRATEO = 6 Then
       
                'Grabamos Movimiento
            GRABAR_MOVIMIENTO "13", Trim(G_CODDELEG), Format(TX_Fech.Text, "YYYYMMDD"), Trim(G_POLINPOL), _
                G_Cero, "OPERAUTO", G_POLINUSU, "''", "", ""
                'Grabamos Suplemento
                ''P415 - JL 03/11/09 - Inicio: Nuevos campos de domicilio
                '' Cargamos concatenados los campos de domicilio en G_POLI2DOMI, ya que la funci�n no admite m�s par�mtros
                G_POLI2DOMI = Trim(G_POLI2_CDG_TIPOVIA) & "#" 'CDG_TIPOVIA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_NOMBREVIA) & "#" 'NOMBREVIA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_NUMEROVIA) & "#" 'NUMEROVIA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_PORTAL) & "#" 'PORTAL
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_BLOQUE) & "#" 'BLOQUE
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_ESCALERA) & "#" 'ESCALERA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_PISO) & "#" 'PISO
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_PUERTA) & "#" 'PUERTA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_RESTOVIA) & "#" 'RESTOVIA
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_CPOBLA_INE) & "#" 'CPOBLA_INE
                G_POLI2DOMI = Trim(G_POLI2DOMI) & Trim(G_POLI2_CVIA_INE) '      CVIA_INE
            
            G_SUPLINDVTAC_SUP = G_POLI2INDVTAC
            G_SUPLINDPRIC_SUP = G_POLI2INDPRIC
                ''P415 - JL 03/11/09 - Fin   : Nuevos campos de domicilio
            INSERTAR_SUPLEMENTO_TCSUPL Trim(G_POLI2CDDE), Trim(G_POLI2NPOL), _
                                   Trim(G_Cero), Trim(G_POLI2NUSU), _
                                   Trim(G_POLI2TIPO), Trim(G_POLI2ESPA), _
                                   "", Trim(G_POLI2FECB), _
                                   Trim(G_POLI2FECA), Trim(G_POLI2FEBA), Trim(G_POLI2CDPT), _
                                   Trim(G_POLI2CDTA), Trim(G_POLI2FOPA), _
                                   Trim(G_POLI2TIPA), Trim(G_POLI2NUPE), _
                                   Trim(G_POLI2IDCP), Trim(G_POLI2IDCO), _
                                   Trim(G_POLI2CDTR), Trim(G_POLI2IDEX), Trim(G_POLI2COBR), _
                                   Trim(G_POLI2AGTA), Trim(G_POLI2AGTB), G_CAT_TIPO_ANT, Trim(G_POLI2PRNT), _
                                   Trim(G_POLI2PRNE), Trim(G_POLI2RECA), Trim(G_POLI2RECE), Trim(G_POLI2IMPT), _
                                   Trim(G_POLI2IMPE), Trim(G_POLI2TORE), Trim(G_POLI2MOBA), Trim(G_POLI2DOMI), _
                                   Trim(G_POLI2CDPS), Trim(G_POLI2CDPO), Trim(G_POLI2TFNO), "", _
                                   "", "", _
                                   "", "", _
                                   "", "", _
                                   Trim(G_POLI2IPUN), Trim(G_POLI2IPUE), _
                                   "", "", _
                                   "", "", "", "", _
                                   sFECM, strSuplemento_DTSUAS, G_POLI2PROVINCIA_TARIFICACION, _
                                   G_POLI2SWPROVINCIA, G_POLI2SWPRODUCCION, "", "", G_POLI2DCTONUMPERSONA, _
                                   G_POLI2RECFORMAPAGO, G_GruposDescuento() 'G_POLI2TIPO_DCTO
  
            F_STAT = "UPDATE DTPOLI SET "
            F_STAT = F_STAT & " POLIFECM = " & Format(Now, "YYYYMMDD")
            F_STAT = F_STAT & " ,POLINUSU= " & Trim(Val(G_POLINUSU) + 1)
            F_STAT = F_STAT + " WHERE POLICDDE = "
            F_STAT = F_STAT + G_POLICDDE + " And POLINPOL = " + G_POLINPOL
            HSTMT = SQL_EXEC(G_HDBC, F_STAT, 0)
            F_ENDCODE = SQL_END(HSTMT)
            Call PMS_BORRA_TMPROR(G_POLINPOL, "0")
            ihstmt = SQL_COMMIT(G_HDBC)
            
            ''''INTEGRACION CON BDI'''''''''''''''''''''''''''''''''''''
            Dim sError As String
            If Not IntegrarConBDI_ODBC(G_HDBC, cteTipoProcesoPoliza, "635", "", Trim(G_POLINPOL), "", sError) Then
                MsgBox "No se pudo enviar la Integraci�n de la P�liza a BDI", vbCritical, "ATENCION"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        ElseIf G_TECLA_PRORRATEO = 7 Then
             lHSTMT = SQL_ROLLBACK(G_HDBC)
        End If
        If G_TECLA_PRORRATEO <> 2 Then
            G_SW_CAMBIO = 0
                    
            Unload FrmFechaPror
            G_SW_CAMBIO = 0
            G_SW_SUPLEMENTO = 0
            G_Sw_Salir = 0
            G_SW_SALVADO_P = 0
            Unload DCAS015
            Unload DCAS017
                    
            G_TECLA_PRORRATEO = 0
            G_SW_RETARIFICAR = False
            DCAS001.Show        
        End If
    End If
End Sub


''
'' STARIND.BAS
''
Public Sub PMS_TARMTTOIND()
    
    Dim a As String
    Dim b As String
    Dim c As String
    Dim D As String
    Dim Fecha_Tarificacion As String
    Dim FECHA_OPERACION As String
    Dim Forma_Pago As String
    Dim ANOO As String
    Dim F_DIAS As Integer
    Dim FA_DIAS As Integer
    Dim FN_DIAS As Integer
    Dim FU_DIAS As Integer  ' D�AS IMPUESTO ANUAL
    Dim FUA_DIAS As Integer  ' D�AS IMPUESTO ANUAL ANTIGUO
    Dim FUN_DIAS As Integer  ' D�AS IMPUESTO ANUAL NUEVO
    Dim P_FECHA_DIAS As String * 10
    Dim F_TARIFI As String
    Dim FechaMovimiento As String
    Dim Usuario As String
    Dim FECHA_TAAN As String
    Dim M_P_TORE_aux As Double
    Dim lhstmt_DTPOCL As Long
    Dim iEnd_DTPOCL As Integer
    Dim aDTPOCL() As String * 255
    Dim bCambio_TipoDescuento As Boolean
    Dim F_POLIFECA As String
    Dim strUsuario          As String
    Dim bProvincia          As Boolean
    Dim bEmitir_Prorrateo   As Boolean
    Dim bOtros_Suplementos  As Boolean
    Dim UltEmision          As String
    Dim Mes                 As Integer
    Dim A�O                 As Integer
    Dim strFechaHasta       As String
    Dim Fecha_Efecto        As String
    Dim strMovimiento       As String
    Dim strValorOld1        As String
    Dim strValorOld2        As String
    Dim strCurrTomadorNIF   As String               ' Se obtendr� el NIF del tomador de la p�liza actual.
    Dim p_codRamo           As String
    Dim p_codModa           As String
    Dim bCambioTomador      As Boolean                                    
    Dim dblPrimaInicial As Double
    Dim dblPrimaCalc As Double
    Dim dblPrimaFinal As Double
    Dim blnMantPricVC As Boolean
    Dim strPricVC As String
    Dim dblTechoSuelo As Double
    Dim intTrataTechos  As Integer
    Dim intTrataSuelos  As Integer
    Dim strAsegInic As String
    Dim blnAplicaTecho As Boolean
    Dim blnOptimizacion As Boolean
    Dim strInforAsegurados As String
    Dim strSQL As String
    Dim lhstmt As Long
    Dim lFETCHCODE As Long
    Dim lENDCODE As Long
    Dim strCom As String
    Dim blnCambia As Boolean
    Dim strDatoPet As String
    Dim sDtoPet(0 To 7) As SubDescuento
    Dim sDtoCol(0) As SubDescuento 'Es necesario para la llamada la funci�n, pero no se cargar�. No necesita redimensioarse a todos los descuentos
    Dim dblSuelo As Double
    Dim dblPrimaFinalSuelo As Double

    Usuario = UCase$(Environ$("USR"))
    If StringVacio(SwProrrateos, "N") = "N" Then
        bEmitir_Prorrateo = False
    Else
        bEmitir_Prorrateo = True
    End If
    
    
    Screen.MousePointer = 11
    
    bCambioTomador = CBool(Trim(DCAS015.TX_Cliecdcl(0).Text) <> G_POLI_TOMADOR)
    
    strCurrTomadorNIF = Trim(DCAS015.TX_Clienif(0).Text)  ' Recuperamos el NIF del tomador del campo de DCAS015... no nos fiemos de las var. globales.
    p_codRamo = Mid(G_POLICDPT, 1, 2)
    p_codModa = Mid(G_POLICDPT, 3, 2)
    strUsuario = UCase$(Environ$("USR"))
    
    bEmitir_Prorrateo = Not CBool(StringVacio(SwProrrateos, "N") = "N")
    
    Forma_Pago = Mid$(Trim$(DCAS015.CB_Fopa.Text), 71, 2)

    G_SW_TARMTTO = 0
    'G_SW_RETARIFICAR es una variable booleana que toma el valor TRUE cuando se pulsa la opci�n RETARIFICAR
    
    ' Comprobaci�n de modificaci�n de los campos que afectan a la tarificaci�n
    
    If G_Sw_Abrir = -1 And G_Sw_Cob_Esp = -1 Then
        G_SW_TARMTTO = 1
    ElseIf G_Sw_Abrir = -1 And G_SWSUPL = -1 Then
        G_SW_TARMTTO = 1
    ElseIf G_POLI2IDEX <> G_CHECK Then
        G_SW_TARMTTO = 1
    ElseIf Trim(G_POLI2CDTA) <> Trim(DCAS015.TX_CDTA.Text) Then
        G_SW_TARMTTO = 1
    ElseIf Trim(G_POLI2FVTAR) <> Trim(G_POLIFVTAR) Then
        G_SW_TARMTTO = 1
    ElseIf G_POLI2FOPA <> Forma_Pago Then
        G_SW_TARMTTO = 1
    ElseIf G_POLI2TIPA <> Mid$(Trim$(DCAS015.CB_Tipa.Text), 71, 2) Then
        G_SW_TARMTTO = 1
    ElseIf Trim$(G_POLI2FECM) <> Trim$(G_POLIFECM) Then
        G_SW_TARMTTO = 1
    ElseIf Trim$(G_POLI2FECB) <> Trim$(DCAS015.TX_Fech(2).Text) Then
        G_SW_TARMTTO = 1
    'p207 Un cambio en la provincia de tarificaci�n produce una retarificaci�n.
    ElseIf Trim(G_POLI2PROVINCIA_TARIFICACION) <> Mid$(Trim$(DCAS017.Cmb_ProvinciaTar), 71, 2) Then
            G_SW_TARMTTO = 1
    ElseIf CambioDescuento(G_GruposDescuento, G_Grupos2Descuento) Then  'JSS
        G_SW_TARMTTO = 1
    ElseIf G_Grupos2Descuento(4).Valor <> G_GruposDescuento(4).Valor Then    'JSS
        'El descuento de retenci�n ha podido cambiar manualmente
        G_SW_TARMTTO = 1
    'P614.50 Control de cambio de valor de descuento comercial
    ElseIf G_Grupos2Descuento(3).Valor <> G_GruposDescuento(3).Valor Then    'JSS
        'El descuento de comercial ha podido cambiar manualmente
        G_SW_TARMTTO = 1
    ElseIf G_POLI2ESPA <> Mid$(Trim$(DCAS015.CB_Espa.Text), 71, 2) Then
        G_SW_TARMTTO = 1
    End If
    
    If G_SW_RETARIFICAR Then G_SW_TARMTTO = 1
    
    'Recuperaci�n de los datos antiguos de DTPOLI
    M_STAT = "SELECT POLICDDE, POLINPOL, POLIFECA, POLICDPT,"
    M_STAT = M_STAT + " POLIFOPA, POLIIDCO, POLICOBR, POLIPRNT, POLIPRNE,"
    M_STAT = M_STAT + " POLIRECA, POLIRECE, POLIIMPT, POLIIMPE, POLITORE,"
    M_STAT = M_STAT + " POLIIPUN, POLIIPUE, POLINUPE, POLIFEER, null POLITIPO_DCTO, POLITREN, POLIFEVE " 'JSS politipo_dcto desaparece
    M_STAT = M_STAT + " FROM DTPOLI "
    M_STAT = M_STAT + " WHERE POLICDDE = " + Trim$(G_POLICDDE) + " AND POLINPOL = " + Trim$(G_POLINPOL)
    M_STAT = M_STAT + " FOR UPDATE"
    
    M_HSTMT3 = SQL_EXEC(G_HDBC, M_STAT, 0, False, "TCPOLI")
    M_FETCHCODE = SQL_FETCH(M_HSTMT3, "Next", MIARRAY())

    If M_FETCHCODE = 0 Then
        F_POLICDDE = RTrim$(MIARRAY(0))
        G_POLICDDE = F_POLICDDE
        F_POLINPOL = Trim$(MIARRAY(1))
        G_POLINPOL = F_POLINPOL
        F_POLIFECA = RTrim$(MIARRAY(2))
        If Trim$(F_POLIFECA) <> "" Then F_POLIFECA = GEN_CTOD(RTrim$(MIARRAY(2)))
        
        G_POLIFECA = F_POLIFECA
        F_POLICDPT = RTrim$(MIARRAY(3))
        G_POLICDPT = F_POLICDPT
        F_POLIFOPA = RTrim$(MIARRAY(4))
        G_POLIFOPA = F_POLIFOPA
        F_POLIIDCO = RTrim$(MIARRAY(5))
        G_POLIIDCO = F_POLIIDCO
        F_POLICOBR = RTrim$(MIARRAY(6))
        G_POLICOBR = F_POLICOBR
        f_poliprnt = CCurNoNulls(RTrim$(MIARRAY(7)))
        G_POLIPRNT = f_poliprnt
        f_poliprne = CCurNoNulls(RTrim$(MIARRAY(8)))
        G_POLIPRNE = f_poliprne
        F_POLIRECA = RTrim$(MIARRAY(9))
        G_POLIRECA = F_POLIRECA
        F_POLIRECE = RTrim$(MIARRAY(10))
        G_POLIRECE = F_POLIRECE
        f_poliimpt = CCurNoNulls(RTrim$(MIARRAY(11)))
        G_POLIIMPT = f_poliimpt
        f_poliimpe = CCurNoNulls(RTrim$(MIARRAY(12)))
        G_POLIIMPE = f_poliimpe
        f_politore = CCurNoNulls(RTrim$(MIARRAY(13)))
        G_POLITORE = f_politore
        f_poliipun = CCurNoNulls(RTrim$(MIARRAY(14)))
        G_POLIIPUN = f_poliipun
        f_poliipue = CCurNoNulls(RTrim$(MIARRAY(15)))
        G_POLIIPUE = f_poliipue
        F_POLINUPE = Trim$(MIARRAY(16))
        G_POLINUPE = F_POLINUPE
        F_POLIFEER = Trim$(MIARRAY(17))
        G_POLIFEER = F_POLIFEER
        G_POLITIPO_DCTO = Trim$(MIARRAY(18))
        G_POLITREN = Trim$(MIARRAY(19))
        G_POLIFEVE = Trim$(MIARRAY(20))
    End If

    M_ENDCODE = SQL_END(M_HSTMT3)

    If G_TECLA_PRORRATEO = 0 Then G_TECLA_PRORRATEO = 6
                    
    'Si ha habido alg�n cambio entra a retarificar
    If G_SW_TARMTTO = 1 Or (G_SW_TARMTTO = 0 And bCambioTomador) Then
            'Comprueba el suplemento realizado
            bCambio_FPago = (Trim$(G_POLI2FOPA) <> Trim$(Forma_Pago))
            bCambio_Tarifa = (Trim(G_POLI2CDTA) <> Trim(DCAS015.TX_CDTA.Text)) Or ((Trim(G_POLI2FVTAR) <> Trim(G_POLIFVTAR) And G_POLI2ESPA <> "04"))
            bReactivacion = (Trim$(G_POLI2ESPA) <> G_ALTAP And Mid$(DCAS015.CB_Espa.Text, 71, 2) = G_ALTAP)
                        
            bCambio_TipoDescuento = CambioDescuento(G_GruposDescuento, G_Grupos2Descuento)  'JSS
            
            bProvincia = (Trim$(G_POLI2PROVINCIA_TARIFICACION) <> Trim$(Mid(DCAS017.Cmb_ProvinciaTar.Text, 71, 2)))
            bOtros_Suplementos = (Not bCambio_FPago And Not bCambio_Tarifa And Not bReactivacion And Not bCambio_TipoDescuento And Not bProvincia)
            If (bCambio_FPago And (bCambio_Tarifa Or bCambio_TipoDescuento Or bReactivacion Or bProvincia)) Or _
                (bCambio_Tarifa And (bReactivacion Or bProvincia)) Or _
                (bCambio_TipoDescuento And (bReactivacion Or bProvincia)) Or (bReactivacion And bProvincia) Or _
                (Not bOtros_Suplementos And G_SWSUPL = -1) Then
                    MsgBox "Realice un s�lo cambio a la vez", vbCritical, "Modificaciones Incorrecta"
                    G_TECLA_PRORRATEO = vbCancel
                Exit Sub
            End If
            
            If bCambio_Tarifa Then
              'Se pierde el descuento si es de pricing, no si es de venta cruzada
              If G_GruposDescuento(5).CODIGO <> G_Descuento_Vacio Or G_GruposDescuento(6).CODIGO = G_Descuento_PricingGEN Then
                iRetorno = MsgBox("Realizar un cambio de tarifa puede implicar perder los descuentos actuales, incluyendo siniestralidad y pricing." & vbCrLf & vbCrLf & _
                  "Si desea mantener los descuentos de siniestralidad y pricing, pulse SI." & vbCrLf & vbCrLf & _
                  "Si desea quitar los descuentos de siniestralidad y pricing, pulse NO." & vbCrLf & vbCrLf & _
                   "Si desea revisar antes el resto de descuentos y modificarlos manualmente, pulse CANCELAR", vbYesNoCancel)
              Else
                iRetorno = MsgBox("Realizar un cambio de tarifa puede implicar perder los descuentos actuales." & vbCrLf & vbCrLf & _
                  "Si desea quitar estos descuentos manualmente, pulse NO." & vbCrLf & vbCrLf & _
                  "Si desea continuar con los descuentos definidos, pulse SI.", vbYesNo) * 10 'Para diferenciar las dos respuestas sobre la misma variable
              End If
              If iRetorno = vbCancel Or iRetorno = vbNo * 10 Then
                G_TECLA_PRORRATEO = vbCancel
                Exit Sub
              Else
                If iRetorno = vbNo Then
                  If G_GruposDescuento(5).CODIGO <> G_Descuento_Vacio Or G_GruposDescuento(6).CODIGO = G_Descuento_PricingGEN Then
                    G_GruposDescuento(5).CODIGO = G_Descuento_Vacio
                    G_GruposDescuento(5).Valor = 0
                    If G_GruposDescuento(6).CODIGO = G_Descuento_PricingGEN Then
                      G_GruposDescuento(6).CODIGO = G_Descuento_Vacio
                      G_GruposDescuento(6).Valor = 0
                    End If
                    bCambio_TipoDescuento = True
                    SalvarDescuentos G_POLINPOL, 0, GEN_DTOC2(Now), G_GruposDescuento
                  End If
                End If
              End If
            End If
        If (bCambio_Tarifa Or bCambio_TipoDescuento Or bProvincia) And Not bReactivacion Then
            Fecha_Correcta = True
            
            strFecha_Operacion = InputBox("Introduzca una fecha de efecto para generar el prorrateo", "FECHA DE EFECTO", G_POLIFECA)
            
            If Not IsDate(Trim(strFecha_Operacion)) Then
                Fecha_Correcta = False
            Else
                strFecha_Operacion = Format(strFecha_Operacion, "DD/MM/YYYY")
                If GEN_DTOC2(G_POLIFECA) > GEN_DTOC2(strFecha_Operacion) Or GEN_DTOC2(strFecha_Operacion) > G_POLIFEVE Then
                    Fecha_Correcta = False
                Else
                    If GEN_DTOC2(strFecha_Operacion) > Format(Now, "YYYYMMDD") And strFecha_Operacion <> Trim$(G_POLIFECA) Then
                        Fecha_Correcta = False
                    End If
                End If
            End If
                
            If Not Fecha_Correcta Then
                strValorOld1 = Format(Now, "YYYYMMDD")
                If GEN_DTOC2(strFecha_Operacion) > G_POLIFEVE Then
                  If GEN_DTOC2(strFecha_Operacion) < Format(Now, "YYYYMMDD") Then
                    strValorOld1 = G_POLIFEVE
                  End If
                End If
                MsgBox "La fecha debe estar comprendida entre el " & G_POLIFECA & " y el " & GEN_CTOD(strValorOld1), vbCritical, "Fecha Incorrecta"
                G_TECLA_PRORRATEO = vbCancel
                Exit Sub
            End If
            
        Else
            If bCambio_FPago Then
                'Comprueba si las formas de pago antigua y nueva ten�an la misma fecha de �ltima emision
                If Proxima_Emision(Trim$(Forma_Pago), G_POLIFEER, Trim(G_POLITREN), , , G_POLIFEVE) <> Proxima_Emision(Trim$(G_POLI2FOPA), G_POLIFEER, Trim(G_POLITREN), , , G_POLIFEVE) Then
                  'Se comprueba si la fecha de alta es mayor que la �ltima cartera emitida. Esto podr�a ocurrir con p�lizas con fecha de alta a futuro y que se les hace un cambio
                  'de forma de pago antes de estar activas.
                  If Left(GEN_DTOC2(G_POLIFECA), 6) > Left$(UltEmision_General, 6) Then
                    strFecha_Operacion = G_POLIFECA
                    If day(strFecha_Operacion) = 1 Then
                        'Se pasa como fecha hasta un mes antes a la fecha de alta para que no se emitan prorrateos.
                        strFechaHasta = G_POLIFECA
                        strFechaHasta = Format(DateAdd("d", -1, strFechaHasta), "yyyymm")
                    Else
                      'Se emitir�n prorrateos hasta la fecha que corresponda
                      strFechaHasta = Format(DateAdd("d", -1, Proxima_Emision(Trim$(G_POLIFOPA), GEN_DTOC2(strFecha_Operacion), Trim(G_POLITREN), "dd/mm/yyyy", , G_POLIFEVE)), "yyyymm")
                    End If
                  Else
                    'P�lizas ya activas.
                    'Comprueba si el cambio de forma de pago ha sido de una forma de pago 'menor' a una 'mayor'. Es decir, si la cantidad de meses que se cobra en un recibo
                    'en la forma de pago antigua es mayor que la cantidad de meses que se cobra en un recibo en la forma de pago nueva.
                    If Meses_FPago(Trim$(Forma_Pago)) < Meses_FPago(Trim$(G_POLI2FOPA)) Then
                      UltEmision = UltEmision_Poliza(Trim$(G_POLINPOL), "yyyymm")
                      A�O = Left(UltEmision, 4) - 1
                      Mes = Right(UltEmision, 2)
                      Mes = Mes + 12 - Meses_FPago(Trim$(G_POLI2FOPA))
                      If Mes >= 12 Then
                        A�O = A�O + 1
                        Mes = Mes - 12
                      End If
                      strFecha_Operacion = "01" & "/" & Format(Mes + 1, "00") & "/" & A�O
                      'El �ltimo mes correspondiente a la forma de pago nueva se calcula desde la fecha de inicio hasta la pr�xima emisi�n.
                      strFechaHasta = Format$(DateAdd("d", -1, strFecha_Operacion), "yyyymm")
                      strFechaHasta = Format$(Proxima_Emision(Trim$(Forma_Pago), strFechaHasta, Trim(G_POLITREN), "dd/mm/yyyy", , G_POLIFEVE), "yyyymm")
                    Else
                      'Si la cantidad de meses a cobrar en un recibo en la nueva forma de pago es mayor que en la forma antigua
                      'La fecha de inicio que se considera para el c�lculo de prorrateo ser� el d�a siguiente a la �ltima emisi�n realizada en el certificado.
                      strFecha_Operacion = DateAdd("d", 1, UltEmision_Poliza(Trim$(G_POLINPOL), "dd/mm/yyyy"))
                      'La fecha hasta ser� la de la pr�xima emisi�n del certificado con la nueva forma de pago, partiendo de la fecha calculada anteriormente.
                      strFechaHasta = Format$(Proxima_Emision(Trim$(Forma_Pago), GEN_DTOC2(strFecha_Operacion), Trim(G_POLITREN), "dd/mm/yyyy", , G_POLIFEVE), "yyyymm")
                    End If
                    strFechaHasta = Left(GEN_DTOC2(DateAdd("d", -1, GEN_CTOD(strFechaHasta & "01"))), 6)
                  End If

                  'Actualiza la fecha de la �ltima emisi�n en funci�n de la nueva forma de pago y de la fecha de la �ltima cartera solicitada
                  lhstmt = SQL_EXEC(G_HDBC, "UPDATE DTPOLI SET POLIFEER = '" & strFechaHasta & "' WHERE POLINPOL = " & Trim$(G_POLINPOL), 0)
                  iEnd = SQL_END(lhstmt)
                  
                  'mira si existen prorrateos pendientes
                  lhstmt_DTPOCL = SQL_EXEC(G_HDBC, "SELECT " & _
                                               "PRORNPOL, PRORCDCE " & _
                                               "FROM STPROR " & _
                                               "WHERE " & _
                                               "PRORNPOL = " & Trim$(G_POLINPOL) & " AND " & _
                                               "PRORCDCE = 0  AND " & _
                                               "PRORSITU = '01' AND PRORTORE <> PRORIPUN + PRORIPUE", 0)
                  'Mueve el cursor al primer registro
                  If SQL_FETCH(lhstmt_DTPOCL, "Next", aDTPOCL()) = 0 Then
                      MsgBox "Compruebe los prorrateos pendientes de la p�liza"
                  End If
                  'Cierra el cursor
                  iEnd_DTPOCL = SQL_END(lhstmt_DTPOCL)
                  Erase aDTPOCL
                Else
                  MsgBox "Las formas de pago son similares"
                  G_TECLA_PRORRATEO = vbYes
                  Exit Sub
                End If
            Else
                If G_Fecha_Asegurado <> "" Then
                    If G_POLI2ESPA <> G_ALTAP And Mid$(DCAS015.CB_Espa.Text, 71, 2) = G_ALTAP Then
                        'Operaciones de REHABILITACION
                        strFecha_Operacion = Trim$(G_POLIFECA)
                        MsgBox "Revise los prorrateos y recibos previos"
                    Else
                        'Operaciones de INCLUSION o EXCLUSION de asegurados
                        strFecha_Operacion = Trim$(G_Fecha_Asegurado)
                    End If
                Else
                    strFecha_Operacion = G_POLIFECM
                End If
            End If
        End If
                
        F_TARIFI = strFecha_Operacion
        
        'Comprueba si el efecto de la operaci�n es postdatado
        bPostDatado = (CDate(F_TARIFI) > CDate(GEN_CTOD(Trim$(F_POLIFEER) & day(DateSerial(Left$(F_POLIFEER, 4), Right(F_POLIFEER, 2) + 1, 0)))))
        
        'Comprueba si se ha realizado alg�n cambio sobre la forma de pago
        If bCambio_FPago Then
            SDTARIFI.Pasa_Parametros Cte_FormaPago, _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(G_POLINPOL), _
                            0, _
                            Trim$(G_POLINUPE), _
                            "", "", Trim(DCAS015.TX_CDTA.Text), _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(Forma_Pago), _
                            Trim(G_POLIPROVINCIA_TARIFICACION), _
                            ConvierteGruposDescuento(G_GruposDescuento), _
                            ConvierteValoresDescuento(G_GruposDescuento), _
                            "", _
                            "", _
                            "", _
                            "S", _
                            CDbl(f_poliprnt), _
                            CDbl(f_poliprne), _
                            CDbl(f_poliimpt), _
                            CDbl(f_poliimpe), _
                            CDbl(f_poliipun), _
                            CDbl(f_poliipue), _
                            CDbl(f_politore)
                            
                        lhstmt = SQL_EXEC(G_HDBC, "UPDATE DTPOLI SET POLIFEER = '" & strFechaHasta & "' WHERE POLINPOL = " & Trim$(G_POLINPOL), 0)
                        iEnd = SQL_END(lhstmt)
        End If
        'Comprueba si se ha realizado alg�n cambio de tarifa
        If bCambio_Tarifa Then
        
            If bCambio_FPago Then
                'Si se ha realizado la operaci�n anterior y se realiza esta en el mismo momento, hay que volver a cargar las primas en TMPROR
                Call PMS_BORRA_TMPROR(Trim$(G_POLINPOL))
                'Carga las primas actuales como 'antigua'
                CARGAR_PRIMAS_ANTIGUAS_POLINDIVIDUALES Trim$(G_POLINPOL) 'Carga las primas actuales como 'antigua'
            End If
        
            If (Trim(G_POLI2CDTA) <> Trim(DCAS015.TX_CDTA.Text)) Then
               'Actualiza la fecha de la de versi�n a 01/01/yyyy
               G_POLIFVTAR = Left(Format$(Now, "YYYYMMDD"), 4) & "0101"
                lhstmt = SQL_EXEC(G_HDBC, "UPDATE DTPOLI SET POLIFVTAR = '" & G_POLIFVTAR & "' WHERE POLINPOL = " & Trim$(G_POLINPOL), 0)
            Else
                'Actualiza la fecha de versi�n a la seleccionada
                lhstmt = SQL_EXEC(G_HDBC, "UPDATE DTPOLI SET POLIFVTAR = '" & G_POLIFVTAR & "' WHERE POLINPOL = " & Trim$(G_POLINPOL), 0)
            End If
            iEnd = SQL_END(lhstmt)

            '' p207 varOLD contiene la tarifa antigua
            SDTARIFI.Pasa_Parametros Cte_CambioTarifa, _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(G_POLINPOL), _
                            0, _
                            Trim$(G_POLINUPE), _
                            Trim(G_POLI2CDTA), G_POLI2FVTAR, Trim(DCAS015.TX_CDTA.Text), _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(Forma_Pago), _
                            Trim(G_POLIPROVINCIA_TARIFICACION), _
                            ConvierteGruposDescuento(G_GruposDescuento), _
                            ConvierteValoresDescuento(G_GruposDescuento), _
                            "", _
                            "", _
                            "", _
                            "S", _
                            CDbl(f_poliprnt), _
                            CDbl(f_poliprne), _
                            CDbl(f_poliimpt), _
                            CDbl(f_poliimpe), _
                            CDbl(f_poliipun), _
                            CDbl(f_poliipue), _
                            CDbl(f_politore)
                            

        End If
        'Comprueba si se ha est� realizando una reactivaci�n
        If bReactivacion Then
            'para la realizaci�n de suplementos de reactivaci�n, se deben tener en cuenta dos premisas muy importantes.
            'Solamente se pueden realizar reactivaciones con fecha de efecto igual a la fecha de alta anterior o con fecha
            'de efecto posterior a la fecha de baja anterior.
            'Retarifica la p�liza con fecha de efecto igual al d�a siguiente a la fecha de baja
            'Tarifica la p�liza con la nueva tarifa, indicando que las primas antiguas son '0'
            'P207 varOLD contiene la fecha de baja inicial.
            
            'Si es una reactivaci�n sobre una anulaci�n sin efecto eliminamos el registro correspondiente en t_gen_histtari
            If G_POLI2FECB = G_POLI2FECA And G_POLIFECA <> G_POLI2FECA And G_POLIFEVE <> G_POLI2FEVE Then
              M_STAT = "DELETE T_GEN_HISTTARI WHERE HISTNPOL = " & Trim(G_POLINPOL) & " AND HISTCDCE = 0"
              M_STAT = M_STAT & " and histfech <'" & G_POLIFEVE & "'"
              M_HSTMT = SQL_EXEC(G_HDBC, M_STAT, 0)
              ENDCODE = SQL_END(M_HSTMT)
            End If
            
            blnCambia = False
            'Pone en estado caducado peticiones automaticas antiguas
            lhstmt = SQL_EXEC(G_HDBC, "UPDATE PETICION_AUT set Estado = '04' WHERE ESTADO= '01' AND FECHA_PETICION<'" & GEN_DTOC2(DateAdd("m", -12, Now)) & "'", 0)
            iEnd = SQL_END(lhstmt)
              
            strDatoPet = ""
            ObtenerTarifaReactivacion Trim$(G_POLINPOL), 0, G_POLIFETA, strDatoPet
              
            If strDatoPet <> "" Then
              If Trim(G_POLICDTA) <> strDatoPet Then
                'P614.50 Se quita mensaje 'La tarifa de la p�liza es distinta ... y se asume que el usuario siempre pulsa la opcion 'SI', cambiar la tarifa
                G_POLICDTA = strDatoPet
                blnCambia = True
              End If
            End If
            
            If blnCambia Then
              
              'Hay que actualizar aqu� la fecha de versi�n para que la funci�n Pasa_par�metros se ejecute con el valor correcto.
              'En una reactivaci�n la versi�n que se usar� ser� la primera del a�o de la fecha de tarificaci�n calculada.
              G_POLIFVTAR = Left(G_POLIFETA, 4) & "0101"
              
              lhstmt = SQL_EXEC(G_HDBC, "UPDATE DTPOLI SET POLIFVTAR = '" & Trim(G_POLIFVTAR) & "' WHERE POLINPOL = " & G_POLINPOL, 0)
              iEnd = SQL_END(lhstmt)
              
              DCAS015.TX_CDTA.Text = G_POLICDTA
        
              'Actualizamos la petici�n Autom�tica
              lhstmt = SQL_EXEC(G_HDBC, "UPDATE PETICION_AUT SET ESTADO = '02'" & _
                  "WHERE POLIZA = " + Trim(G_POLINPOL) & " AND CERTIFICADO = 0 AND TIPO_PETICION = '01'" & _
                  " AND FECHA_PETICION <= '" & G_POLIFETA & "' AND TO_CHAR(ADD_MONTHS(TO_DATE(FECHA_PETICION, 'YYYYMMDD'), 12), 'YYYYMMDD') > '" & G_POLIFETA & "'", 0)
              iEnd = SQL_END(lhstmt)
              
            End If
            
            blnCambia = False
            'P614.50 Correcci�n fechas, misma correcci�n que en certificados.
            ObtenerTipoDescuentoReactivacion Trim$(G_POLINPOL), 0, G_POLIFETA, sDtoPet()
            
            If ExisteDescuento(sDtoPet, False) Then
              'No se compara el dcto de venta cruzada o pricing porque el usuario no cambia este tipo de descuento
              sDtoPet(6) = G_GruposDescuento(6)
              If CambioDescuento(sDtoPet, G_GruposDescuento) Then
                strCom = ""
                lhstmt = SQL_EXEC(G_HDBC, "SELECT * FROM T_DES_GRUPOS WHERE DGRUCODG < 90 ORDER BY 1", 0)
                ' Mueve el cursor al primer registro
                lFETCHCODE = SQL_FETCH(lhstmt, "Next", MIARRAY())
                For i = 0 To UBound(sDtoPet)
                  If i + 1 = CInt(MIARRAY(0)) And i <> 6 Then
                    strCom = strCom & Trim(MIARRAY(1)) & String(1 + IIf(i = 3 Or i = 4, 1, 0), vbTab) & String(7, " ") & G_GruposDescuento(i).CODIGO & String(2, vbTab) & String(4, " ") & IIf(sDtoPet(i).CODIGO = c_DescNoFiltrar, G_GruposDescuento(i).CODIGO, sDtoPet(i).CODIGO) & vbCrLf
                  End If
                  lFETCHCODE = SQL_FETCH(lhstmt, "Next", MIARRAY())
                Next i
                ' Cierra el cursor y libera recursos
                iEnd = SQL_END(lhstmt)
                'P614.50 Se quita mensaje 'El descuento de la p�liza es distinto ... y se asume que el usuario siempre pulsa la opcion 'SI', cambiar el descuento
                AsignarArrayGrupoDescuentos G_GruposDescuento, sDtoPet
                blnCambia = True
                lhstmt = SQL_EXEC(G_HDBC, "UPDATE PETICION_AUT SET ESTADO = '02'" & _
                    "WHERE POLIZA = " + Trim(G_POLINPOL) & " AND CERTIFICADO = 0 AND TIPO_PETICION = '02'" & _
                    " AND FECHA_PETICION <= '" & G_POLIFETA & "' AND TO_CHAR(ADD_MONTHS(TO_DATE(FECHA_PETICION, 'YYYYMMDD'), 12), 'YYYYMMDD') > '" & G_POLIFETA & "'", 0)
                iEnd = SQL_END(lhstmt)
              End If 'CambioDescuento(sDtoPet, G_GruposDescuento)
            End If 'If ExisteDescuento(sDtoPet, False)
            
            'Se ha decidido que no se tratar� un nuevo descuento de pricing en reactivaciones posteriores a la renovaci�n (16/05/2017). Se mantiene el c�digo por ahora
            'Si ha cambiado el periodo se vuelven a recuperar los descuentos
            'If G_POLIFETA > G_POLI2FETA Then ObtenerGruposDescuentoCertificado G_POLINPOL, 0, G_POLIFETA, G_GruposDescuento, blnIniciaPricing:=True
            
            'Tratamiento para calcular descuento de siniestralidad
            blnOptimizacion = False
            blnAplicaTecho = False
            If G_GruposDescuento(5).CODIGO <> "000" And G_GruposDescuento(5).CODIGO <> G_Descuento_Optimizacion Then
              dblPrimaInicial = 0
              strAsegInic = ""
              strInforAsegurados = ""
              'Recoge el porcentaje de siniestralidad para el c�lculo fijo
              'P614.50 Calculamos el porcentaje del c�digo de descuento, ya sea nuevo o el mismo. No aplicamos la suma ya que no procede, s�lo en renovaci�n.
              G_GruposDescuento(5).Valor = CalculaPorcentajeGrupo(G_GruposDescuento(5).CODIGO, G_POLIFETA, "G" & Trim(G_POLIVIP), dblSuelo)
              'Comprueba si hay pricing. Si hay pricing se activa la variable para poner el c�digo de descuento al finalizar el tratamiento de siniestralidad
              blnMantPricVC = False
              If G_GruposDescuento(6).CODIGO <> "000" Then
                blnMantPricVC = True
                strPricVC = G_GruposDescuento(6).CODIGO
              End If
              'Quita el pricing
              G_GruposDescuento(6).CODIGO = "000"
              'Recoge el porcentaje de siniestralidad para techo/suelo
              dblTechoSuelo = CalculaPorcentajeGrupo(obtenerDescEdad(G_GruposDescuento(5).CODIGO, G_POLIFETA), G_POLIFETA, "G" & Trim(G_POLIVIP), dblSuelo)
              'Seg�n el resultado que se obtenga posteriormente con los techos y suelos habr� que volver a tarificar o no. Por ello se mete en un bucle
              intTrataTechos = 0 'Si no hay techos/suelos no es necesario hacer la retarificaci�n. Se va con el valor de porcentaje fijo
              intTrataSuelos = 0 'Si no hay techos/suelos no es necesario hacer la retarificaci�n. Se va con el valor de porcentaje fijo
              'P614.50 Se comprueba el suelo y techo.
              If dblTechoSuelo <> 0 Or dblSuelo <> 0 Then
                intTrataTechos = 1
                intTrataSuelos = 1
                'Hay que obtener los asegurados que van a estar de alta en la renovaci�n y que est�n de alta actualmente
                strSQL = "SELECT POCLCDCL, SUM(POPCIPTP) FROM DTPOCL "
                strSQL = strSQL & " INNER JOIN TSPOPC ON POCLNPOL = POPCNPOL AND POCLCDCE = POPCCDCE AND POCLCDCL = POPCCDCL"
                strSQL = strSQL & " WHERE POCLNPOL = " & G_POLINPOL & " AND POCLCDRE <> '01' "
                strSQL = strSQL & "AND NVL(POCLFECB,'99999999') >= '" & G_POLIFETA & "' AND POCLFECA <  '" & G_POLIFETA & "' GROUP BY POCLCDCL"
                lhstmt = SQL_EXEC(G_HDBC, strSQL, 0)
                lFETCHCODE = SQL_FETCH(lhstmt, "Next", MIARRAY())
                Do While (lFETCHCODE = 0)
                  strAsegInic = strAsegInic & Trim(MIARRAY(0)) & ", "
                  dblPrimaInicial = dblPrimaInicial + CDbl(Trim(MIARRAY(1)))
                  lFETCHCODE = SQL_FETCH(lhstmt, "Next", MIARRAY())
                Loop
                lENDCODE = SQL_END(lhstmt)
                 'El tratamiento en PL/SQL a�ade comilla simple al valor que se pasa, por lo que quitamos el primer car�cter y los tres �ltimos
                  strAsegInic = Left(strAsegInic, Len(strAsegInic) - 2)
                  dblPrimaFinal = dblPrimaInicial * dblTechoSuelo / 100 'El valor de techo viene expresado como porcentaje total
                  dblPrimaFinalSuelo = dblPrimaInicial * dblSuelo / 100 'El valor de suelo viene expresado como porcentaje total

                strInforAsegurados = strAsegInic
                Do While intTrataTechos <> 0 Or intTrataSuelos <> 0
                  'Se calcula la prima sin el techo y sin el pricing
                  Pasa_Parametros Cte_TarificacionOnline, _
                     G_POLIFETA, _
                     G_POLINPOL, _
                     0, _
                     G_POLINUPE, _
                     Trim(G_POLIFVTAR), "", _
                     Trim(G_POLICDTA), _
                     G_POLIFETA, _
                     G_POLIFOPA, _
                     G_POLIPROVINCIA_TARIFICACION, _
                     ConvierteGruposDescuento(G_GruposDescuento), _
                     ConvierteValoresDescuento(G_GruposDescuento), _
                     "", _
                     "", _
                     "", _
                     "", _
                     "0", "0", "0", _
                     "0", "0", "0", _
                     "0", strInforAsegurados
                
                  dblPrimaCalc = CDbl(G_POLIPRNT) + CDbl(G_POLIPRNE)
                  'P614.50 Nuevo algoritmo de c�lculo de suelos y techos.
                     If dblPrimaCalc > 0 Then
                      If intTrataTechos = 1 Or intTrataSuelos = 1 Then 'Primera vez para tratar techos/suelos
                          If dblPrimaCalc > dblPrimaFinal Then
                            intTrataTechos = 2
                            G_GruposDescuento(5).Valor = 0
                            strInforAsegurados = strAsegInic
                          Else
                            intTrataTechos = 0
                          End If
                          If dblPrimaCalc < dblPrimaFinalSuelo Then
                            intTrataSuelos = 2
                            G_GruposDescuento(5).Valor = 0
                            strInforAsegurados = strAsegInic
                          Else
                            intTrataSuelos = 0
                          End If
                      ElseIf intTrataTechos = 2 Then
                          intTrataTechos = 0
                          blnAplicaTecho = True
                          G_GruposDescuento(5).Valor = Round(-(dblPrimaCalc - dblPrimaFinal) * 100 / dblPrimaCalc, 2)
                      ElseIf intTrataSuelos = 2 Then
                          intTrataSuelos = 0
                          blnAplicaTecho = True
                          G_GruposDescuento(5).Valor = Round(-(dblPrimaCalc - dblPrimaFinalSuelo) * 100 / dblPrimaCalc, 2)
                      End If
                    Else
                      intTrataTechos = 0
                      intTrataSuelos = 0
                    End If
                  Loop
                'Restablece el c�digo de VC o Pricing
                If blnMantPricVC Then G_GruposDescuento(6).CODIGO = strPricVC
              End If
            Else
              'P614.50 Tratamiento de descuento de optimizaci�n para la reactivaci�n.
              If G_GruposDescuento(5).CODIGO = G_Descuento_Optimizacion Then
                  blnOptimizacion = True
              End If
            End If 'If G_GruposDescuento(5).Codigo <> "000"
            
            'Se vac�a la variable para que haga el tratamiento a todos los asegurados que corresponden
            strInforAsegurados = ""
            
            'Se ha decidido que no se tratar� un nuevo descuento de pricing en reactivaciones posteriores a la renovaci�n (16/05/2017)
            'En reactivaciones posteriores a la renovaci�n, se quitar� el descuento de venta cruzada o pricing
            If G_GruposDescuento(6).CODIGO <> "000" And G_POLIFETA <> G_POLI2FETA Then
              G_GruposDescuento(6).CODIGO = "000"
              G_GruposDescuento(6).Valor = "000"
              blnCambia = True
            End If
            
            If blnOptimizacion Then
              'Actualiza el estado
              M_STAT = "UPDATE T_DES_OPTIMIZACION SET OPTIESTA = 'R' "
              M_STAT = M_STAT & " WHERE OPTINPOL = " & G_POLINPOL
              M_STAT = M_STAT & " AND OPTICDCE = 0 AND OPTIANOMES = '" & Left(G_POLIFETA, 6) & "'"
              
              lhstmt = SQL_EXEC(G_HDBC, M_STAT, 0)
              ENDCODE = SQL_END(lhstmt)
            End If

            strCom = ""
            If G_POLIFETA <> G_POLI2FETA Then
              
              M_STAT = "UPDATE T_SIN_GRUPOS SET GRUPESTA = 'R' "
              M_STAT = M_STAT & " WHERE GRUPNPOL = " & G_POLINPOL
              M_STAT = M_STAT & " AND GRUPCDCE = 0 AND GRUPANOMES = '" & Left(G_POLIFETA, 6) & "'"
              
              lhstmt = SQL_EXEC(G_HDBC, M_STAT, 0)
              ENDCODE = SQL_END(lhstmt)
              
              'Corresponde a una reactivaci�n despu�s de la renovaci�n
              strCom = " POLIVIP_ANT = '" & G_POLIVIP & "', "
              strCom = strCom & " POLIVIP = " & IIf(G_POLIGSIN = "", "NULL", "'" & Right(G_POLIGSIN, 1) & "'") & ","
              strCom = strCom & " POLIGSIN = NULL, "
            End If
            
            M_STAT = "UPDATE DTPOLI SET " & strCom
            M_STAT = M_STAT & " POLIINDTESU = '" & IIf(blnAplicaTecho, "S", "N") & "'"
            M_STAT = M_STAT & " , POLIDESC_G06 = " & Grabacion_Importe(G_GruposDescuento(5).Valor, True)
            M_STAT = M_STAT & " , POLIDESC_G07 = " & Grabacion_Importe(G_GruposDescuento(6).Valor, True)
            M_STAT = M_STAT & " , POLIINDOPT = '" & IIf(blnOptimizacion, "S", "N") & "'"
            M_STAT = M_STAT & " WHERE POLICDDE = " & G_POLICDDE
            M_STAT = M_STAT & " AND POLINPOL = " & G_POLINPOL
            
            lhstmt = SQL_EXEC(G_HDBC, M_STAT, 0)
            lENDCODE = SQL_END(lhstmt)

            If G_POLI2FECA <> G_POLIFECA Then
              lhstmt = SQL_EXEC(G_HDBC, "UPDATE T_DES_POLIZAS_DCTO SET DPODFECH = '" & GEN_DTOC2(G_POLIFECA) & "' WHERE DPODNPOL = " & G_POLINPOL & _
                " AND DPODCDCE = 0 AND DPODFECH = '" & GEN_DTOC2(G_POLI2FECA) & "'", 0)
              iEnd = SQL_END(lhstmt)
            End If
            
            If blnCambia Then SalvarDescuentos Trim$(G_POLINPOL), 0, IIf(G_POLIFETA > GEN_DTOC2(G_POLIFECA), G_POLIFETA, GEN_DTOC2(G_POLIFECA)), G_GruposDescuento
            
            SDTARIFI.Pasa_Parametros Cte_Reactivacion, _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(G_POLINPOL), _
                            0, _
                            Trim$(G_POLINUPE), _
                            Trim(GEN_DTOC2(G_POLI2FECB)), "", Trim(DCAS015.TX_CDTA.Text), _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(Forma_Pago), _
                            Trim(G_POLIPROVINCIA_TARIFICACION), _
                            ConvierteGruposDescuento(G_GruposDescuento), _
                            ConvierteValoresDescuento(G_GruposDescuento), _
                            "", _
                            "", _
                            "", _
                            "S", _
                            "0", _
                            "0", _
                            "0", _
                            "0", _
                            "0", _
                            "0", _
                            "0"
                            
        End If

        If bOtros_Suplementos Or bCambio_TipoDescuento Or bProvincia Then                                             
            If bCambio_FPago Or bCambio_Tarifa Then
                'Si se ha realizado la operaci�n anterior y se realiza esta en el mismo momento, hay que volver a cargar las primas en TMPROR
                Call PMS_BORRA_TMPROR(Trim$(G_POLINPOL))
                'Carga las primas actuales como 'antigua'
                CARGAR_PRIMAS_ANTIGUAS_POLINDIVIDUALES Trim$(G_POLINPOL) 'Carga las primas actuales como 'antigua'
            End If
                                             
            'En el caso de que no se haya encontrado una tarifa para la fecha de efecto del suplemento, pero si el d�a de efecto es superior a la fecha de vencimiento o inferior a
            'la fecha de tarificaci�n, se graba el suplemento pero no se genera prorrateo
            bEmitir_Prorrateo = True
            If CDate(F_TARIFI) > DateAdd("d", 1, GEN_CTOD(G_POLIFEVE)) Or (GEN_DTOC2(F_TARIFI) < G_POLIFETA And Not (Val(G_POLINUPE) > Val(G_POLI2NUPE))) Then
                MsgBox "Se ha encontrado que para la fecha de efecto del suplemento no hay definidas unas primas." & vbCr & vbCr & _
                  "Debido a esto, no se emitir� prorrateo. En caso de querer emitir prorrateo tendr� que realizarse de forma manual.", vbInformation, "Suplementos de certificado"
                bEmitir_Prorrateo = False
            End If
                        
            strValorOld1 = Trim(G_POLI2FECB)
            
            ' La validaci�n se realiza solamente si se a�aden asegurados. En el caso de exclusiones no se vuelven a recalcular los descuentos.
            If bOtros_Suplementos And Val(G_POLINUPE) > Val(G_POLI2NUPE) Then
               strMovimiento = Cte_InclusionOnLine  ' Es una inclusion
            Else
              If bProvincia Then
                strMovimiento = Cte_CambioProvincia
                strValorOld1 = G_POLI2PROVINCIA_TARIFICACION
              Else
                strMovimiento = Cte_OtrosSuplementosOnline
              End If
            End If

            G_POLI_CIA = DCAS015.usc_Companias1.Compania    ' TODO: Hay que poner mejor esto de estas 3 "G_*" de abajo...
                                                                                                                                                                                                         
            Dim strCurrCodAgente As String
            strCurrCodAgente = Trim(DCAS017.TX_AgteNif(0).Text)     ' El c�digo del agente principal.
                                                                                                                                                                                             
            ' Una manera de manejar la situaci�n cuando se est� Re-Tarificando... en el marco de un mantenimiento
            ' de p�liza, provocamos que la fecha de alta de la PI sea la misma que la de Tarificaci�n.
            
            If G_Sw_Abrir = -1 Then
                If G_SW_RETARIFICAR Then
                  Debug.Print "G_SW_RETARIFICAR And G_Sw_Abrir = -1 -> Se evita Query al WS. G06= " + _
                              MdlTiposDescuento.G_GruposDescuento(6).CODIGO & "/" & MdlTiposDescuento.G_GruposDescuento(6).Valor
                End If
            Else
                MdlTiposDescuento.Query_PricingVCWS Trim(G_CODDELEG), Trim$(G_POLINPOL), strCurrTomadorNIF, Trim(G_POLI_CIA), p_codRamo, p_codModa, "000", "VC", F_TARIFI, strUsuario, "0", G_Sw_Abrir, G_POLIFETA, G_POLIFEVE, strCurrCodAgente, cTiposAccion.RegaloDirClientes, G_Sw_Salir                       '       el m�dulo debe ser mas general, est� hecho para globales
                Debug.Print G_POLVTAC_MENSJ
                If MdlTiposDescuento.QRYPRICINGWS_STR <> "" Then MsgBox MdlTiposDescuento.QRYPRICINGWS_STR, vbCritical, "ERROR AL INSERTAR"
            End If
                                        
            If GEN_DTOC2(DCAS015.TX_Fech(2).Text) = G_POLIFEVE Then
                bEmitir_Prorrateo = False
            Else
                bEmitir_Prorrateo = True
            End If
            If G_POLI2ESPA = G_ALTAP And Mid(DCAS015.CB_Espa, 71, 2) = G_BAJA_POLIZA Then
                strValorOld1 = IIf(bEmitir_Prorrateo, "S", "N")
                strValorOld2 = Trim(Mid(DCAS015.CB_Mobp, 71, 2))
                strMovimiento = Cte_BajaOnline
                F_POLIFECA = G_POLIFECA
                FechaMovimiento = Trim(DCAS015.TX_Fech(2).Text)
                FechaMovimiento = Format$(DateAdd("d", 1, FechaMovimiento), "DD/MM/YYYY")
            Else
                strValorOld1 = strValorOld1
                strValorOld2 = G_POLI2MOBA
                strMovimiento = strMovimiento
                F_POLIFECA = F_TARIFI
                FechaMovimiento = F_TARIFI
            End If
            Pasa_Parametros strMovimiento, _
                            GEN_DTOC2(FechaMovimiento), _
                            Trim$(G_POLINPOL), _
                            0, _
                            Trim$(G_POLINUPE), _
                            strValorOld1, Trim(G_POLI2MOBA), Trim(DCAS015.TX_CDTA.Text), _
                            GEN_DTOC2(F_TARIFI), _
                            Trim$(Forma_Pago), _
                            Trim(Mid(DCAS017.Cmb_ProvinciaTar.Text, 71, 2)), _
                            ConvierteGruposDescuento(G_GruposDescuento), _
                            ConvierteValoresDescuento(G_GruposDescuento), _
                            "", _
                            "", _
                            "", _
                            IIf(bEmitir_Prorrateo, "S", "N"), _
                            CDbl(f_poliprnt), _
                            CDbl(f_poliprne), _
                            CDbl(f_poliimpt), _
                            CDbl(f_poliimpe), _
                            CDbl(f_poliipun), _
                            CDbl(f_poliipue), _
                            CDbl(f_politore)
                            
        End If
        
        'Muestra la pantalla con la informaci�n de los prorrateos y primas anules resultantes
        Mostrar_Pantalla_Confirmacion Trim$(G_POLINPOL), _
                                      Trim$(G_POLIPRNT), _
                                      Trim$(G_POLIPRNE), _
                                      Trim$(G_POLIIMPT), _
                                      Trim$(G_POLIIMPE), _
                                      Trim$(G_POLIIPUN), _
                                      Trim$(G_POLIIPUE), _
                                      Trim$(G_POLIRECA), _
                                      Trim$(G_POLIRECE), _
                                      Trim$(G_POLITORE), _
                                      Trim$(F_POLIFEER), _
                                      Trim$(Forma_Pago), _
                                      GEN_DTOC2(strFecha_Operacion), _
                                      bPostDatado, GEN_DTOC2(F_TARIFI)
        
        Call PMS_BORRA_DETALLE_PRORRATEO(Trim$(G_POLINPOL), "0", "")    ' Borra los datos para el c�lculo de los prorrateos
    End If
    
End Sub


''
'' IBERCAJA.bas
''
Public Sub GRABAR_MOVIMIENTO(sTipo As String, sDelegacion As String, sFecha As String, sPoliza As String, sCertificado As String, sUsuario As String, sSuplemento As String, sCliente As String, sAntes As String, sDespues As String, Optional sFecha_Efecto = "", Optional sNumExpSGO As String = "")
'Procedimiento que inserta un registro en la tabla TCMOVI.

    Dim Stat As String
    Dim sql_Fields As String
    Dim sql_Values As String
    
    Dim HSTMT As Long
    Dim ENDCODE As Integer
    
    sql_Fields = "": sql_Values = ""
    sql_Fields = "MOVINUME,": sql_Values = "SQ_MOVIM.NEXTVAL,"
    sql_Fields = sql_Fields & " MOVITIPO,": sql_Values = sql_Values & "'" & Trim(sTipo) & "',"
    sql_Fields = sql_Fields & " MOVICDDE,": sql_Values = sql_Values & Trim(sDelegacion) & ","
    sql_Fields = sql_Fields & " MOVIFECH,": sql_Values = sql_Values & "'" & Trim(sFecha) & "',"
    sql_Fields = sql_Fields & " MOVINPOL,": sql_Values = sql_Values & Trim(sPoliza) & ","
    sql_Fields = sql_Fields & " MOVICDCE,": sql_Values = sql_Values & Trim(sCertificado) & ","
    sql_Fields = sql_Fields & " MOVIUSUA,": sql_Values = sql_Values & "'" & Trim(sUsuario) & "',"
    sql_Fields = sql_Fields & " MOVINUSU,": sql_Values = sql_Values & Trim(sSuplemento) & ","
    sql_Fields = sql_Fields & " MOVICDCL"
    If Trim(sCliente) = "" Then
      sql_Values = sql_Values & "NULL"
    Else
      sql_Values = sql_Values & Trim(sCliente)
    End If
    If UCase(sAntes) <> "NULL" And NoNulls(sAntes) <> "" Then
        sql_Fields = sql_Fields & ", MOVIANTE": sql_Values = sql_Values & ",'" & Trim(sAntes) & "'"
    End If
    If UCase(sDespues) <> "NULL" And NoNulls(sDespues) <> "" Then
        sql_Fields = sql_Fields & ", MOVIDESP":  sql_Values = sql_Values & ",'" & Trim(sDespues) & "'"
    End If
    If NoNulls(sFecha_Efecto) <> "" Then
        sql_Fields = sql_Fields & ", MOVIFECH_EFECTO":  sql_Values = sql_Values & ",'" & Trim(sFecha_Efecto) & "'"
    End If
    If NoNulls(sNumExpSGO) <> "" Then
        sql_Fields = sql_Fields & ", MOVITAREA_SGO":  sql_Values = sql_Values & ",'" & Trim(sNumExpSGO) & "'"
    End If
    Stat = "INSERT INTO TCMOVI(" & sql_Fields & ") VALUES (" & sql_Values & ")"
    
    HSTMT = SQL_EXEC(G_HDBC, Stat, 0)
    ENDCODE = SQL_END(HSTMT)

End Sub

''
'' mdlSuplementos.bas
''
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
                           ByVal G_POCE2IPUN As String, ByVal G_POCE2IPUE As String, ByVal G_POCE2CPUN As String, ByVal G_POCE2CPUE As String, ByVal G_Fecha_Asegurado As String, ByVal G_POLI2IDMA As String, ByVal G_POLI2NUCE As String, ByVal G_POLI2CDRP As String, _
                           ByRef G_POCEFECM As String, ByRef strSuplemento_DTSUAS As String, _
                           ByVal G_POCEPROVINCIA_TARIFICACION As String, ByVal G_POCESWPROVINCIA As String, ByVal G_POCESWPRODUCCION As String, _
                           ByVal G_POCESWTARIFA As String, ByVal G_POCESWDCTO As String, ByVal G_POCEDCTONUMPERSONAS As String, ByVal G_POCERECFORMAPAGO As String, _
                           G_Descuentos() As SubDescuento) 'No se pueden poner m�s par�metros en la definici�n de la funci�n. Se usan variables globales en este m�dulo

   Dim strSql As String
   Dim ihstmt_DTSUPL As Long
   Dim iEnd As Integer
   
   'INSERTA SUPLEMENTO DE CERTIFICADO
   Screen.MousePointer = 11
   
   If Trim$(G_Fecha_Asegurado) <> "" Then
      G_POCEFECM = G_Fecha_Asegurado
   Else
      G_POCEFECM = Format(Now, "dd/mm/yyyy")
   End If
    
   strSql = "INSERT INTO TCSUPL (SUPLCDDE, SUPLNPOL, SUPLCDCE, SUPLNUSU,"
   strSql = strSql & " SUPLTIPO, SUPLSITP, SUPLSITC, SUPLFECA, SUPLFECB, SUPLFECC,"
   strSql = strSql & " SUPLFEBA, SUPLCDPT, SUPLCDTA, SUPLFOPA, SUPLTIPA,"
   strSql = strSql & " SUPLNUPE, SUPLIDCP, SUPLIDCO, SUPLCDTR, SUPLIDEX,"
   strSql = strSql & " SUPLCOBR, SUPLAGTA, SUPLAGTB, SUPLINSP, SUPLPRNT,"
   strSql = strSql & " SUPLPRNE, SUPLRECA, SUPLRECE, SUPLIMPT, SUPLIMPE,"
   ''P415 - JL 03/11/09 se elimina de grabar el campo SUPLDOMI y se graba en los nuevos datos de domicilio
   'strSQL = strSQL & " SUPLTORE, SUPLMOCE, SUPLDOMI, SUPLCDPS, SUPLCDPO,"
   strSql = strSql & " SUPLTORE, SUPLMOCE, SUPLCDPS, SUPLCDPO,"
   strSql = strSql & " SUPLTFNO_NUTE, SUPLCRNT, SUPLCRNE, SUPLCECA, SUPLCECE,"
   strSql = strSql & " SUPLCMPT, SUPLCMPE, SUPLCORE, SUPLIPUN, SUPLIPUE,"
   strSql = strSql & " SUPLCPUN, SUPLCPUE, SUPLIDMA, SUPLNUCE, SUPLCDRP, SUPLCDPC, "
   'P207 MT
   strSql = strSql & " SUPLPROVINCIA_TARIFICACION,SUPLSWPROVINCIA,SUPLSWPRODUCCION,"
   strSql = strSql & " SUPLSWTARIFA,SUPLSWDCTO,SUPLDCTONUMPERSONAS,SUPLRECFORMAPAGO,"
   
   'JSS Eliminamos SUPLTIPO_DCTO
   strSql = strSql & " SUPLADAP, SUPLTERRITORIALIDAD,"
   'strSQL = strSQL & " SUPLTIPO_DCTO, SUPLADAP, SUPLTERRITORIALIDAD,"
   'JSS
   
   ''P415 - JL 03/11/09 - Inicio: Nuevos campos de domicilio
   If InStr(G_POCE2DOMI, "#") = 0 Then ' solo viene relleno el  campo Domi con toda la direcci�n antigua
        strSql = strSql & " SUP_NOMBREVIA"
   Else
        strSql = strSql & " SUP_CDG_TIPOVIA, SUP_NOMBREVIA, SUP_NUMEROVIA,"
        strSql = strSql & " SUP_PORTAL, SUP_BLOQUE, SUP_ESCALERA,"
        strSql = strSql & " SUP_PISO, SUP_PUERTA, SUP_RESTOVIA,"
        strSql = strSql & " SUP_CPOBLA_INE, SUP_CVIA_INE"
   End If
   ''P415 - JL 03/11/09 - Fin   : Nuevos campos de domicilio
   
   'JSS Nuevos campos de grupos de descuentos
   strSql = strSql & ", SUPLDESC_G01, SUPLDESC_G02, SUPLDESC_G03, SUPLDESC_G04, SUPLDESC_G05, SUPLDESC_G06, SUPLDESC_G07 , SUPLDESC_G08 "
   'JSS
   
   strSql = strSql & ", SUPLINDVTAC, SUPLINDPRIC "   ' [614.30.R.IU.VB.282] MJBF: Se persistir� a partir de ahora tambi�n el valor del campo "Aplica VC" en TCSUPL.

   strSql = strSql & " ) VALUES ("
   strSql = strSql & Grabar_BBDD(G_POCE2CDDE, "N")                        'SUPLCDDE
   strSql = strSql & "," & Grabar_BBDD(G_POCE2NPOL, "N")                   'SUPLNPOL
   strSql = strSql & "," & Grabar_BBDD(G_POCE2CDCE, "N")                   'SUPLCDCE
   strSql = strSql & "," & Grabar_BBDD(G_POCE2NUSU, "N")                   'SUPLNUSU
   strSql = strSql & "," & Grabar_BBDD(G_CER, "S")                         'SUPLTIPO
   strSql = strSql & "," & Grabar_BBDD(G_POLI2ESPA, "S")             'SUPLSITP
   strSql = strSql & "," & Grabar_BBDD(G_POCE2ALBA, "S")             'SUPLSITC
   strSql = strSql & ",'" & Format(Now, "YYYYMMDD") & "'"      'SUPLFECA
   strSql = strSql & "," & GEN_DTOC1(Trim(G_POCE2FECB))               'SUPLFECB
   strSql = strSql & "," & GEN_DTOC1(Trim(G_POCE2FECA))               'SUPLFECC
   strSql = strSql & "," & GEN_DTOC1(Trim(G_POCE2FEBA))               'SUPLFEBA
   strSql = strSql & "," & Grabar_BBDD(G_POLI2CDPT, "N")                   'SUPLCDPT
   strSql = strSql & "," & Grabar_BBDD(G_POCE2CDTA, "S")             'SUPLCDTA
   strSql = strSql & "," & Grabar_BBDD(G_POCE2FOPA, "S")              'SUPLFOPA
   strSql = strSql & "," & Grabar_BBDD(G_POCE2TIPA, "S")              'SUPLTIPA
   strSql = strSql & "," & Grabar_BBDD(G_POCE2NUPE, "N")              'SUPLNUPE
   strSql = strSql & "," & Grabar_BBDD(G_POLI2IDCP, "S")              'SUPLIDCP
   strSql = strSql & "," & Grabar_BBDD(G_POLI2IDCO, "S")              'SUPLIDCO
   strSql = strSql & "," & Grabar_BBDD(G_POLI2CDTR, "S")              'SUPLCDTR
   strSql = strSql & "," & Grabar_BBDD(G_POLI2IDEX, "S")              'SUPLIDEX
   strSql = strSql & "," & Grabar_BBDD(UCase(G_POCE2COBR), "S")     'SUPLCOBR
   strSql = strSql & "," & Grabar_BBDD(UCase(G_POCE2AGTA), "S")     'SUPLAGTA
   strSql = strSql & "," & Grabar_BBDD(UCase(G_POCE2AGTB), "S")     'SUPLAGTB
   strSql = strSql & "," & Grabar_BBDD(UCase(G_CAT_TIPO_ANT), "S")  'SUPLINSP
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2PRNT), True)                  'SUPLPRNT
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2PRNE), True)                  'SUPLPRNE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2RECA), True)                  'SUPLRECA
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2RECE), True)                  'SUPLRECE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2IMPT), True)                  'SUPLIMPT
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2IMPE), True)                  'SUPLIMPE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2TORE), True)                'SUPLTORE
   strSql = strSql & "," & Grabar_BBDD(G_POCE2MOBA, "S")             'SUPLMOCE
   ''P415 - JL 03/11/09 se elimina de grabar el campo SUPLDOMI y se graba en los nuevos datos de domicilio
   'strSQL = strSQL & "," & Grabar_BBDD(UCase(G_POCE2DOMI), "S")     'SUPLDOMI
   strSql = strSql & "," & Grabar_BBDD(G_POCE2CDPS, "S")             'SUPLCDPS
   strSql = strSql & "," & Grabar_BBDD(G_POCE2CDPO, "S")             'SUPLCDPO
   RegLog "Graba suplemento G_POCE2CDPO = " & G_POCE2CDPO
   strSql = strSql & "," & Grabar_BBDD(G_POCE2TFNO, "S")              'SUPLTFNO
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CRNT), True)                   'SUPLCRNT
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CRNE), True)                   'SUPLCRNE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CECA), True)                  'SUPLCECA
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CECE), True)                   'SUPLCECE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CMPT), True)                  'SUPLCMPT
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CMPE), True)                'SUPLCMPE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CORE), True)                  'SUPLCORE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2IPUN), True)              'SUPLIPUN
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2IPUE), True)                  'SUPLIPUE
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CPUN), True)                  'SUPLCPUN
   strSql = strSql & "," & Grabacion_Importe(Trim(G_POCE2CPUE), True)                  'SUPLCPUE
   strSql = strSql & "," & Grabar_BBDD(G_POLI2IDMA, "S")              'SUPLIDMA
   strSql = strSql & "," & Grabar_BBDD(G_POLI2NUCE, "N")                    'SUPLNUCE
   strSql = strSql & "," & Grabar_BBDD(G_POLI2CDRP, "N")                    'SUPLCDPR
   strSql = strSql & "," & Grabar_BBDD(G_POLI2CDRP, "N")                    'SUPLCDPR
   'P207
   strSql = strSql & ",'" & Trim$(G_POCEPROVINCIA_TARIFICACION) & "'"
   strSql = strSql & ",'" & Trim$(G_POCESWPROVINCIA) & "'"
   strSql = strSql & ",'" & Trim$(G_POCESWPRODUCCION) & "'"
   strSql = strSql & ",'" & IIf(Trim$(G_POCESWTARIFA) = "", "N", Trim$(G_POCESWTARIFA)) & "'"
   strSql = strSql & ",'" & IIf(Trim$(G_POCESWDCTO) = "", "N", Trim$(G_POCESWDCTO)) & "'"
   strSql = strSql & "," & Grabacion_Importe(Trim$(G_POCEDCTONUMPERSONAS), True) & ""
   strSql = strSql & "," & Grabacion_Importe(Trim$(G_POCERECFORMAPAGO), True) & ""
   
   'JSS
   'strSQL = strSQL & "," & Grabar_BBDD(Trim$(G_POCETIPO_DCTO), "S")
   'JSS
   
   strSql = strSql & "," & Grabar_BBDD(Trim$(G_SUPLADAP_SUP), "S")
   strSql = strSql & "," & Grabar_BBDD(Trim$(G_SUPLTERRITORIALIDAD_SUP), "S")
   ''P415 - JL 03/11/09 - Inicio: Nuevos campos de domicilio
   If InStr(G_POCE2DOMI, "#") = 0 Then ' solo viene relleno el  campo Domi con toda la dirrecci�n antigua
        strSql = strSql & "," & Grabar_BBDD(ConvierteTextoComillaAmpersandOracle2(Trim$(G_POCE2DOMI)), "S")
   Else
        Dim l_POCE2DOMI As String
        l_POCE2DOMI = G_POCE2DOMI
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'CDG_TIPOVIA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(ConvierteTextoComillaAmpersandOracle2(GEN_QUOTE(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "'", "�")), "S") 'NOMBREVIA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'NUMEROVIA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'PORTAL
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'BLOQUE
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'ESCALERA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'PISO
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'PUERTA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(ConvierteTextoComillaAmpersandOracle2(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1))), "S") 'RESTOVIA
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(Mid(l_POCE2DOMI, 1, InStr(l_POCE2DOMI, "#") - 1)), "S") 'CPOBLA_INE
        l_POCE2DOMI = Mid(l_POCE2DOMI, InStr(l_POCE2DOMI, "#") + 1)
        strSql = strSql & "," & Grabar_BBDD(Trim$(l_POCE2DOMI), "S") 'CVIA_INE
   End If
   ''P415 - JL 03/11/09 - Fin   : Nuevos campos de domicilio

    '''   'SUPLCPUE
    '''   strSQL = strSQL & ")"
   
   'JSS
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(0).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(1).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(2).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(3).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(4).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(5).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(6).Valor, "N"), True)
   strSql = strSql & "," & Grabacion_Importe(Grabar_BBDD(G_Descuentos(7).Valor, "N"), True)
   'JSS
   
   strSql = strSql & ",'" & G_SUPLINDVTAC_SUP & "','" & G_SUPLINDPRIC_SUP & "'"
   
   strSql = strSql & ")"
   ihstmt_DTSUPL = SQL_EXEC(G_HDBC, strSql, 0)
   iEnd = SQL_END(ihstmt_DTSUPL)
   
   If G_CER <> "02" Then
        INSERTAR_SUPLEMENTO_TSSUPC G_POCE2NPOL, G_POCE2CDCE, G_POCE2NUSU  ' Insertar un registro por cada fila de la matriz de TSPOPC
   End If
   
   strSuplemento_DTSUAS = G_POCE2NUSU
   Screen.MousePointer = 0
End Sub

''
'' SDTARIFI.bas
''
Public Sub PMS_BORRA_TMPROR(ByVal Poliza As String, _
                            Optional Certificado As Variant)

    'Procedimiento PL/SQL
    Dim lHSTMT As Long
    Dim iParametro As Integer
    Dim iExec As Integer
    Dim iEnd As Integer
    'Definici�n de los par�metros
    Dim paramTipoSql As Integer
    Dim paramColDef As Long
    Dim paramColScale As Integer
    Dim paramNullable As Integer
    Dim dPoliza As Double
    Dim dCertificado As Double
    dPoliza = CDbl(Poliza)
    If IsMissing(Certificado) Then
        dCertificado = -1
    Else
        If Certificado = "" Then
           dCertificado = -1
        Else
           dCertificado = CDbl(Certificado)
        End If
    End If
     'Prepara el cursor
    lHSTMT = SQL_PREPARE(G_HDBC, "{Call PCK_TARPRO.PMS_BORRA_TMPROR(?,?)}")
    'Prepara los par�metros
    'Poliza
    iParametro = SQLDescribeParam(lHSTMT, 1, paramTipoSql, paramColDef, paramColScale, paramNullable)
    iParametro = SQLBindParameter(lHSTMT, 1, SQL_PARAM_INPUT, SQL_C_DOUBLE, paramTipoSql, paramColDef, paramColScale, dPoliza, paramColDef, SQL_NTS)
    'Certificado
    iParametro = SQLDescribeParam(lHSTMT, 2, paramTipoSql, paramColDef, paramColScale, paramNullable)
    iParametro = SQLBindParameter(lHSTMT, 2, SQL_PARAM_INPUT, SQL_C_DOUBLE, paramTipoSql, paramColDef, paramColScale, dCertificado, paramColDef, SQL_NTS)
     'Ejecuta el procedimiento
    iExec = SQL_EXECUTE(lHSTMT)
    'Cierra el cursor
    iEnd = SQL_END(lHSTMT)

End Sub


''
'' SDGEN1.BAS
''
Function PMF_ERROR(CODERROR As String, Optional Incluir As String) As Integer

    Dim Hstmt9 As Long
    Dim fetchcode As Long
    Dim ENDCODE As Long
    'Esta funcion recibe el codigo de error del mensaje que se quiere
    'displayar lo busca en la base de datos DTMERR.DBF y una vez encontrado
    'ese codigo cogemos su DESCRIPCION (MERRDSER en la base de datos) que sera el
    'mensaje a displayar y TIPOERROR (MERRCDTE en la base de datos) que sera un
    'numero que indica el tipo de mensaje (precaucion,warning,error,informativo)
    'y los botones a displayar en la caja ademas del correspondiente icono.

    'Declaraci�n de variables:


    'Conexi�n con la base de datos
    
    Stat$ = "SELECT MERRCDER, MERRDSER, MERRCDTE, MERRDSCT from DTMERR where MERRCDER = '" + CODERROR + "'"

    ReDim MIARRAY(4) As String * 255
    Hstmt9 = SQL_EXEC(G_HDBC, Stat$, 0)
    fetchcode = SQL_FETCH(Hstmt9, "Next", MIARRAY())
    ENDCODE = SQL_END(Hstmt9)
    
    If fetchcode = 0 Then
       M_DESCRIPCION = GEN_QUOTE(Mid$(MIARRAY(1), 1, 70), Chr$(0), " ")
       M_TIPOERROR = Val(GEN_QUOTE(Mid$(MIARRAY(2), 1, 2), Chr$(0), " "))
       M_MENSAJE = GEN_QUOTE(Mid$(MIARRAY(3), 1, 35), Chr$(0), " ")
    End If

'Selecciono el tipo de mensaje segun el tipoerror (Titulo de la caja)

    If Mid(app.Title, 1, 5) = "CTRLM" Then
        RegLog Trim$(M_DESCRIPCION)
        G_I = 6
    Else
        G_I = MsgBox(Trim$(M_DESCRIPCION) & " " & Incluir, M_TIPOERROR, M_MENSAJE)
        blnExiste = True
    End If

    If G_I = 1 Then      'si se pulsa  OK  le asigno el valor
        
       G_I = 6           'de la tecla SI porque hacen lo mismo
    End If
    'DPC 22/12/2003
    '****************************
    If G_TIPOACC = "R" And (G_I = 1 Or G_I = 6) Then
        If Mid(app.Title, 1, 5) = "CTRLM" Then
            RegLog "LOS CAMBIOS NO TENDRAN EFECTO"
            G_I = 6
        Else
           G_I = MsgBox("LOS CAMBIOS NO TENDRAN EFECTO", vbOKOnly, "ACCESO DE LECTURA")
        End If
       If G_I = 1 Then G_I = 7
   ' Else
   '    G_I = PMF_ERROR("CAI081")
    End If
    '****************************
'Devuelvo el numero de la tecla pulsada

    PMF_ERROR = G_I
End Function


''
'' SDTARIFI.bas:2708
''
Public Sub PMS_BORRA_DETALLE_PRORRATEO(ByVal F_POLIZA As String, ByVal F_CERTIFICADO As String, ByVal F_NUPR As String)
    'Procedimiento PL/SQL
    Dim lHSTMT As Long
    Dim iParametro As Integer
    Dim iExec As Integer
    Dim iEnd As Integer
    'Definici�n de los par�metros
    Dim paramTipoSql As Integer
    Dim paramColDef As Long
    Dim paramColScale As Integer
    Dim paramNullable As Integer
    Dim dPoliza As Double
    Dim dCertificado As Double
    Dim dNumProrrateo As Double
    dPoliza = CDbl(F_POLIZA)
    dCertificado = CDbl(F_CERTIFICADO)
    If Trim(F_NUPR) <> "" Then
       dNumProrrateo = CDbl(F_NUPR)
    Else
       dNumProrrateo = -1
    End If
     'Prepara el cursor
    lHSTMT = SQL_PREPARE(G_HDBC, "{Call PCK_TARPRO.PMS_BORRA_DETALLE_PRORRATEO(?,?,?)}")
    'Prepara los par�metros
    'Poliza
    iParametro = SQLDescribeParam(lHSTMT, 1, paramTipoSql, paramColDef, paramColScale, paramNullable)
    iParametro = SQLBindParameter(lHSTMT, 1, SQL_PARAM_INPUT, SQL_C_DOUBLE, paramTipoSql, paramColDef, paramColScale, dPoliza, paramColDef, SQL_NTS)
    'Certificado
    iParametro = SQLDescribeParam(lHSTMT, 2, paramTipoSql, paramColDef, paramColScale, paramNullable)
    iParametro = SQLBindParameter(lHSTMT, 2, SQL_PARAM_INPUT, SQL_C_DOUBLE, paramTipoSql, paramColDef, paramColScale, dCertificado, paramColDef, SQL_NTS)
    'Numero prorrateo
    iParametro = SQLDescribeParam(lHSTMT, 3, paramTipoSql, paramColDef, paramColScale, paramNullable)
    iParametro = SQLBindParameter(lHSTMT, 3, SQL_PARAM_INPUT, SQL_C_DOUBLE, paramTipoSql, paramColDef, paramColScale, dNumProrrateo, paramColDef, SQL_NTS)
    
     'Ejecuta el procedimiento
    iExec = SQL_EXECUTE(lHSTMT)
    'Cierra el cursor
    iEnd = SQL_END(lHSTMT)

End Sub



''
'' MdlTiposDescuento.bas:849
''
'****************************************************************************
'NOMBRE     : CambioDescuento
'PARAMETROS : array con los siete grupos de descuento de la poliza/certificado
'             array con los siete grupos de descuento odificados
'FUNCION    : Comprueba si en el array de descuentos hay algun descuento definido
'****************************************************************************
Public Function CambioDescuento(aDescuentos1() As SubDescuento, aDescuentos2() As SubDescuento) As Boolean
    Dim I As Integer
    CambioDescuento = False
    For I = 0 To UBound(aDescuentos1)
        'Se comparan los descuentos si ninguno de los dos contienen el c�digo que indica que se mantiene el descuento
        If aDescuentos1(I).Codigo <> aDescuentos2(I).Codigo And Not (aDescuentos1(I).Codigo = c_DescNoFiltrar Or aDescuentos2(I).Codigo = c_DescNoFiltrar) Then
            CambioDescuento = True
            Exit Function
        End If
        If aDescuentos1(I).Grupo = "04" Or aDescuentos1(I).Grupo = "05" Then
          If aDescuentos1(I).Valor <> aDescuentos2(I).Valor Then
            CambioDescuento = True
            Exit Function
          End If
        End If
    Next I
End Function