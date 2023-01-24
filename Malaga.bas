Attribute VB_Name = "Malaga"
'11-01-2019
' Elimino los /br
' Elimino opcion de efectivo o tarjeta.
' Nuevo Aviso.docx

'06-02-2019
' Cambio ext 4 por ext 1
'01-03-2019
' Facturas. No se hace porque falta informacion en el mail

'25-04-2019
' WinHTTP

'02-05-2019
'cambio WinHTTP por XMLHTTP para poder enviar el archivo

'17-05-2019
'soluciono fallo que enviaba la fecha vacia

'05-07-2019
'Modificaciones para nuevas apis GPFD

' 21-10-2019 Añado Avisos Serie 99 y Defectuosas.
' 15-11-2019 Añado: El pago también se podrá realizar mediante ingreso en la c/c ES10 0075 0672 5607 5000 1550

'02-03-2020 malaga
' Se añaden auditores, publicidad y deposito cuentas

'20-04-2020
'Se añade carpetas eoores para que el proceso no de un aviso cada vez que encuentra un fallo en una notificacion
'Se añade la funcion EnviaDeposito para diferenciar el envio de notificaciones tipo 2

'24-06-2020
'Se detecta fallos en tildes cuando se envia a mails diferentes a gmail.
'Creo la funcion SustituyeTildes que cambia las vocales acentuadas por acute;
'La aplico al body en EnvioFacturas

'07-02-2022
'Añado form FrmEnvio

Dim ns As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
Dim Items As Outlook.Items
Dim Item As Object
Dim apikey As String

Public Sub Envia()

Set ns = Application.GetNamespace("MAPI")
Set olFolder = ns.GetDefaultFolder(olFolderOutbox)
Set Items = ns.GetDefaultFolder(olFolderOutbox).Items
Dim vacio As Boolean

vacio = True
apikey = "c38519de71df6f2085046ae390301079"


If Items.Count > 0 Then
    If MsgBox("Se procederá a realizar el envío de Notificaciones. ¿Continuar?", vbInformation + vbDefaultButton2 + vbYesNo) = vbYes Then
        FrmEnvio.Show
        vacio = False
        For Each Item In olFolder.Items
            EnviaUno
        Next
        For Each Item In olFolder.Items
            EnviaUno
        Next
        
    End If
    If Items.Count > 0 Then
        vacio = False
        For Each Item In olFolder.Items
            EnviaUno
        Next
        For Each Item In olFolder.Items
            EnviaUno
        Next
        
    End If
FrmEnvio.Hide
    If Not vacio Then MsgBox "Envío finalizado. Revise la bandeja de salida." & Chr(13) & "También debe revisar la bandeja de Errores por si se hubiera detectado algun fallo durante el envío."
Else
    MsgBox "No hay notificaciones para enviar"
End If

End Sub

Private Function EnviaUno() As Long
On Error GoTo ErrorEnvia

Dim miid As Long
Dim mail As String
Dim factura As String
Dim asunto As String
Dim body As String
Dim importe As String
Dim tipo As String
Dim fecha As Date
Dim body1 As String, body2, bodySoc As String
Dim LOPD1, LOPD2, LOPD3, LOPD4, LOPD5 As String
Dim nombre As String
Dim Final_body As String
Dim CodError As String
Dim marca As Long

'Ruta: <url>/api/nueva_factura.php             'Parámetros (Post)
'apikey: Idenficador único del usuario autorizado.
'codigo: código único identificador de la factura por la Aplicación externa.
'tipo: Tipo de documento enviado.
'email: Dirección del destinatario.
'asunto: Asunto del correo electrónico de notificación del correo.
'cuerpo: Cuerpo del correo electrónico (admite forma html). A este cuerpo se incluye automáticamente los datos de enlace hacia la plataforma de pago,
    'y el pie definido en la plataforma. En caso de no existir, no se envía el email.
'fecha: Fecha de la factura (formato yyyy-mm-dd).
'importe: Importe de la factura (formato nnnn.nn).

factura = ExtraeNumFactura(Item.Subject)
If Not ControlNumFactura(factura) Then
    CodError = "Numero de Documento"
    GoTo ErrorEnvia
End If

FrmEnvio.factura = factura

Select Case AnalizaTipo(factura)
    Case 1: tipo = "1" ' papel o telematicas q seria tipo=4
    Case 5: tipo = "5" ' publicidad
    Case 4: tipo = "6" ' auditores
    Case 2: tipo = "2" ' deposito cuentas
    Case 9: tipo = "9" ' serie 99
    Case Else
End Select

EnviaUno = 0

marca = InStr(1, Item.body, "El pago se")
If marca = 0 Then
    marca = InStr(1, Item.body, "Para cualquier")
    If marca = 0 Then marca = Len(Item.body)
End If

If tipo = "2" Then
    EnviaUno = EnviaDeposito(tipo, marca)
Else
    EnviaUno = EnviaPapel(tipo, marca)
End If

If EnviaUno <> 0 Then
    Item.Move (olFolder.Folders("GPFD"))
Else
    Item.Move (olFolder.Folders("Errores"))
End If
    
Exit Function

ErrorEnvia:
EnviaUno = 0
'MsgBox "No se ha podido enviar la notificación. " & factura & "  Error:" & CodError

End Function

Private Function EnviaPapel(tipo As String, marca As Long) As Long
Dim mail As String
Dim factura As String
Dim asunto As String
Dim body As String
Dim importe As String
Dim fecha As Date
Dim body1 As String, body2, bodySoc As String
Dim LOPD1, LOPD2, LOPD3, LOPD4, LOPD5 As String
Dim nombre As String
Dim Final_body As String
Dim CodError As String

'***************** LOPD para cuerpo de mail de GPFD ********************************************
LOPD1 = "<p><pre>" & "******************** ADVERTENCIA LEGAL ******************** "
LOPD2 = "<br>" & "Este mensaje contiene informaci&oacute;n confidencial destinada para ser leída exclusivamente por el destinatario."
LOPD3 = "<br>" & "Queda prohibida su reproducci&oacute;n, publicaci&oacute;n y divulgaci&oacute;n total o parcial del mensaje, así como el uso no autorizado por el emisor."
LOPD4 = "<br>" & "Si Vd. lo ha recibido por error, le rogamos que por favor lo destruya inmediatamente y se ponga en contacto con nosotros."
LOPD5 = "<br>" & "Su dirección de correo se encuentra recogida en nuestros ficheros con la finalidad de mantener correspondencia electrónica, responder a las consultas por Vd. planteadas y el envío de comunicaciones por diversos medios, incluyendo los electrónicos, entendiéndose que consiente el tratamiento de los citados datos con dicha finalidad."
LOPD6 = "<br>" & "Usted puede ejercitar sus derechos de acceso, rectificación, cancelación y oposición ante REGISTRADORES MERCANTILES DE MALAGA C.B. de acuerdo a lo previsto en Reglamento General de Protección de Datos 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016 y en la Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales"
LOPD7 = " y garantía de los derechos digitales (BOE núm.  294, de 6 de diciembre de 2018)." & "</pre></p>"
            
FrmEnvio.texto = "Enviando Papel..."

mail = ExtraeMail(Item.Recipients(1))
factura = ExtraeNumFactura(Item.Subject)
asunto = Limpia_Comas(Item.Subject)
body = Mid(Item.body, 1, marca - 1)
body = Limpia_Espacios(Limpia_guiones(body))
               
    marca1_0 = InStr(body, ":") '
    marca1_1 = InStr(body, "/")
    marca2 = InStr(body, "PRESENTANTE:")
    marca3 = InStr(body, "SOCIEDAD:")
    marca4 = InStr(body, "PROTOCOLO:")
    marca5 = InStr(body, "AUTORIZANTE:")
    marca6 = InStr(body, "el cual ya ha sido")
    marca7 = InStr(body, "asciende a:") + 11
    marca8 = InStr(body, "Total Registro")
   
    If marca1_1 <> 0 And marca2 <> 0 Then
        bodyini = Mid(body, marca1_1 - 1, Len(body))
        bodyini = Mid(bodyini, 1, marca2 - marca1_1)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    
    bodyini = factura
    
    If marca2 <> 0 And marca3 <> 0 Then
        bodypre = Mid(body, marca2, Len(body))
        bodypre = Mid(bodypre, 1, marca3 - marca2)
        'bodypre = TextoSinAcentos(bodypre)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    If marca3 <> 0 And marca4 <> 0 Then
        bodySoc = Mid(body, marca3, Len(body))
        bodySoc = Mid(bodySoc, 1, marca4 - marca3)
       ' bodySoc = TextoSinAcentos(bodySoc)
    ElseIf marca4 = 0 Then
        If marca5 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca5 - marca3)
            'bodySoc = TextoSinAcentos(bodySoc)
        ElseIf marca6 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca6 - marca3)
        ElseIf marca6 = 0 Then
            If asunto Like "*defecto*" Then
                marca61 = InStr(body, "El importe")
            Else
                marca61 = InStr(body, "Puede pasar")
            End If
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca61 - marca3)
           ' bodySoc = TextoSinAcentos(bodySoc)
        End If
        'GoTo ErrorEnvia
    End If
    If marca4 <> 0 And marca5 <> 0 Then
        bodyProt = Mid(body, marca4, Len(body))
        bodyProt = Mid(bodyProt, 1, marca5 - marca4)
        bodyProt = bodyProt & "<br>"
    Else
        'GoTo ErrorEnvia
    End If
    If marca5 <> 0 And marca6 <> 0 Then
        bodyAut = Mid(body, marca5, Len(body))
        bodyAut = Mid(bodyAut, 1, marca6 - marca5)
        'bodyAut = TextoSinAcentos(bodyAut)
        bodyAut = bodyAut & "<br>"
    ElseIf marca6 = 0 Then
        marca61 = InStr(body, "Puede pasar") ' defectos
        If marca61 <> 0 And marca5 <> 0 Then
            bodyAut = Mid(body, marca5, Len(body))
            bodyAut = Mid(bodyAut, 1, marca61 - marca5)
        
        End If

    Else
        'GoTo ErrorEnvia
    End If
    If marca6 <> 0 And marca7 <> 0 Then
        bodyelcual = Mid(body, marca6, Len(body))
        bodyelcual = Mid(bodyelcual, 1, marca7 - marca6)
    ElseIf marca6 = 0 Then
        marca61 = InStr(body, "Puede pasar") ' defectos
        If marca61 <> 0 Then
            bodyelcual = Mid(body, marca61, Len(body))
            bodyelcual = Mid(bodyelcual, 1, marca7 - marca61)
        End If
        'CodError = "Faltan datos."
        'GoTo ErrorEnvia
    End If
    
    If marca8 <> 0 Then
        bodyasciende = Mid(body, marca8, Len(body))
        marcae = InStr(bodyasciende, "euros")
        bodyasciende = Mid(bodyasciende, 1, marcae + 4)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    restodevolver = Mid(body, marca8, Len(body) - marca8)
    
    marcar1 = InStr(bodyasciende, "ro")
    bodyresto0 = Left(bodyasciende, marcar1 + 1)
    marcar2 = InStr(bodyasciende, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyasciende, 1, marcar2 + 4))
    bodyresto0 = bodyresto0 & ": " & bodyresto2 & " euros"
    
    marca9 = InStr(body, "Total Provisión") - 1
    bodyresto = Right(body, Len(body) - marca9)
    marcar1 = InStr(bodyresto, ":")
    bodyresto1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1 = bodyresto1 & " " & bodyresto2 & " euros"
    
    marca9_1 = InStr(body, "A PAGAR") - 1
    
    bodyresto = Right(body, Len(body) - marca9_1)
    marcar1 = InStr(bodyresto, "R")
    bodyresto1_1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
    
    marca1 = InStr(body, "A PAGAR")
    If marca1 <> 0 Then
        Final_body = Mid(body, marca1, Len(body))
        'body2 = ""
        body2 = "El pago también se podrá realizar mediante ingreso en la c/c ES10 0075 0672 5607 5000 1550." '  "<br>Forma de Pago: Enlace abajo indicado."
        body2 = body2 & "<br><p>Para cualquier consulta o aclaraci&oacute;n sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 952 28 89 35 ext 1," & "<br>" & "en la direcci&oacute;n de correo malaga@registromercantil.org o en www.rmmalaga.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & " " & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7
    Else
        marca1 = InStr(body, "A DEVOLVER")
        
        marca9_1 = InStr(body, "A DEVOLVER") - 1
    
        bodyresto = Right(body, Len(body) - marca9_1)
        marcar1 = InStr(bodyresto, "R")
        bodyresto1_1 = Left(bodyresto, marcar1)
        marcar2 = InStr(bodyresto, "euros")
        bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
        bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
        '********
        Final_body = Mid(body, marca1, Len(body))
        body2 = ""
        body2 = body2 & "<p>Para cualquier consulta o aclaración sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 952 28 89 35 ext 1," & "<br>" & "en la dirección de correo malaga@registromercantil.org o en www.rmmalaga.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & " " & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>" ' "<br>" & restodevolver & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7
    End If
                                                                
    importe = importe & ExtraeNumeros(Final_body)
                
    fecha = Format(Date, "dd/mm/yyyy")
    
    'FrmEnvio.texto = "Generando Documento..."
    
    ControlWord factura, fecha, bodySoc, importe
    
    FrmEnvio.texto = "Enviando a GPFD..."
    
    EnviaPapel = EnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, "c:\GPFD\Aviso1.pdf")
    

Exit Function

ErrorEnvia:
EnviaPapel = 0
'MsgBox "No se ha podido enviar la notificacion. " & factura & "  Error:" & CodError

End Function

Private Function EnviaDeposito(tipo As String, marca As Long) As Long
Dim mail As String
Dim factura As String
Dim asunto As String
Dim body As String
Dim importe As String
Dim fecha As Date
Dim body1 As String, body2, bodySoc As String
Dim LOPD1, LOPD2, LOPD3, LOPD4, LOPD5 As String
Dim nombre As String
Dim Final_body As String
Dim CodError As String

'***************** LOPD para cuerpo de mail de GPFD ********************************************
LOPD1 = "<p><pre>" & "******************** ADVERTENCIA LEGAL ******************** "
LOPD2 = "<br>" & "Este mensaje contiene informaci&oacute;n confidencial destinada para ser leída exclusivamente por el destinatario."
LOPD3 = "<br>" & "Queda prohibida su reproducci&oacute;n, publicaci&oacute;n y divulgaci&oacute;n total o parcial del mensaje, así como el uso no autorizado por el emisor."
LOPD4 = "<br>" & "Si Vd. lo ha recibido por error, le rogamos que por favor lo destruya inmediatamente y se ponga en contacto con nosotros."
LOPD5 = "<br>" & "Su dirección de correo se encuentra recogida en nuestros ficheros con la finalidad de mantener correspondencia electrónica, responder a las consultas por Vd. planteadas y el envío de comunicaciones por diversos medios, incluyendo los electrónicos, entendiéndose que consiente el tratamiento de los citados datos con dicha finalidad."
LOPD6 = "<br>" & "Usted puede ejercitar sus derechos de acceso, rectificación, cancelación y oposición ante REGISTRADORES MERCANTILES DE MALAGA C.B. de acuerdo a lo previsto en Reglamento General de Protección de Datos 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016 y en la Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales"
LOPD7 = " y garantía de los derechos digitales (BOE núm.  294, de 6 de diciembre de 2018)." & "</pre></p>"
             
FrmEnvio.texto = "Enviando Deposito..."

mail = ExtraeMail(Item.Recipients(1))
factura = ExtraeNumFactura(Item.Subject)
asunto = Limpia_Comas(Item.Subject)
body = Mid(Item.body, 1, marca - 1)
body = Limpia_Espacios(Limpia_guiones(body))
               
    marca1_0 = InStr(body, ":") '
    marca1_1 = InStr(body, "/")
    marca2 = InStr(body, "PRESENTANTE:")
    marca3 = InStr(body, "SOCIEDAD:")
    marca6 = InStr(body, "Puede pasar") ' defectos
    marca7 = InStr(body, "El importe asciende")
    marca8 = InStr(body, "Total Registro")
   
     
    If marca2 <> 0 And marca3 <> 0 Then
        bodypre = Mid(body, marca2, Len(body))
        bodypre = Mid(bodypre, 1, marca3 - marca2)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
          
    If marca3 <> 0 And marca6 <> 0 Then
        bodySoc = Mid(body, marca3, Len(body)) 'defectos
        bodySoc = Mid(bodySoc, 1, marca6 - marca3)
        If marca6 <> 0 And marca7 <> 0 Then
            bodypuede = Mid(body, marca6, Len(body))
            bodypuede = Mid(bodyelcual, 1, marca7 - marca6)
            If marca7 <> 0 And marca8 <> 0 Then
                bodyelcual = Mid(body, marca7, Len(body))
                bodyelcual = bodypuede & "<br>" & Mid(bodyelcual, 1, marca8 - marca7)
            ElseIf marca8 = 0 Then
                GoTo ErrorEnvia
            End If
        ElseIf marca7 = 0 Then
             GoTo ErrorEnvia
        End If
 
    ElseIf marca6 = 0 Then
        bodySoc = Mid(body, marca3, Len(body))
        bodySoc = Mid(bodySoc, 1, marca7 - marca3)
        If marca7 <> 0 And marca8 <> 0 Then
            bodyelcual = Mid(body, marca7, Len(body))
            bodyelcual = Mid(bodyelcual, 1, marca8 - marca7)
        ElseIf marca8 = 0 Or marca8 = 0 Then
            GoTo ErrorEnvia
        End If
 
    End If
    

    bodyasciende = Mid(body, marca8, Len(body))
    marcae = InStr(bodyasciende, "euros")
    bodyasciende = Mid(bodyasciende, 1, marcae + 4)

    restodevolver = Mid(body, marca8, Len(body) - marca8)
    
    marcar1 = InStr(bodyasciende, "ro")
    bodyresto0 = Left(bodyasciende, marcar1 + 1)
    marcar2 = InStr(bodyasciende, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyasciende, 1, marcar2 + 4))
    bodyresto0 = bodyresto0 & ": " & bodyresto2 & " euros"
    
    marca9 = InStr(body, "Total Provisión") - 1
    bodyresto = Right(body, Len(body) - marca9)
    marcar1 = InStr(bodyresto, ":")
    bodyresto1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1 = bodyresto1 & " " & bodyresto2 & " euros"
    
    marca9_1 = InStr(body, "A PAGAR") - 1
    
    bodyresto = Right(body, Len(body) - marca9_1)
    marcar1 = InStr(bodyresto, "R")
    bodyresto1_1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
    
    marca1 = InStr(body, "A PAGAR")
    If marca1 <> 0 Then
        Final_body = Mid(body, marca1, Len(body))
        body2 = "El pago también se podrá realizar mediante ingreso en la c/c ES10 0075 0672 5607 5000 1550." '  "<br>Forma de Pago: Enlace abajo indicado."
        body2 = body2 & "<br><p>Para cualquier consulta o aclaraci&oacute;n sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 952 28 89 35 ext 1," & "<br>" & "en la direcci&oacute;n de correo malaga@registromercantil.org o en www.rmmalaga.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & " " & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7
    Else
        marca1 = InStr(body, "A DEVOLVER")
        
        marca9_1 = InStr(body, "A DEVOLVER") - 1
    
        bodyresto = Right(body, Len(body) - marca9_1)
        marcar1 = InStr(bodyresto, "R")
        bodyresto1_1 = Left(bodyresto, marcar1)
        marcar2 = InStr(bodyresto, "euros")
        bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
        bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
        '********
        Final_body = Mid(body, marca1, Len(body))
        body2 = ""
        body2 = body2 & "<p>Para cualquier consulta o aclaraci&oacute;n sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 952 28 89 35 ext 1," & "<br>" & "en la direcci&oacute;n de correo malaga@registromercantil.org o en www.rmmalaga.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & " " & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>" ' "<br>" & restodevolver & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7
    End If
                                                                
    importe = importe & ExtraeNumeros(Final_body)
                
    fecha = Format(Date, "dd/mm/yyyy")
    
    'FrmEnvio.texto = "Generando documento..."
    
    ControlWord factura, fecha, bodySoc, importe
    
    FrmEnvio.texto = "Subiendo a GPFD..."
    
    EnviaDeposito = EnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, "c:\GPFD\Aviso1.pdf")
    

Exit Function

ErrorEnvia:
EnviaDeposito = 0
'MsgBox "No se ha podido enviar la notificacion. " & factura & "  Error:" & CodError

End Function

Private Function EnviaUnoSeg() As Long
On Error GoTo ErrorEnvia

Dim miid As Long
Dim mail As String
Dim factura As String
Dim asunto As String
Dim body As String
Dim importe As String
Dim tipo As String
Dim fecha As Date
Dim body1 As String, body2, bodySoc As String
Dim LOPD1, LOPD2, LOPD3, LOPD4, LOPD5 As String
Dim nombre As String
Dim Final_body As String
Dim CodError As String

Dim myAttachments As Outlook.Attachments
Dim Fichero1, Fichero2 As String
Dim Num_Adj As Integer
Dim EnvioDoble As Boolean

'Ruta: <url>/api/nueva_factura.php             'Parámetros (Post)
'apikey: Idenficador único del usuario autorizado.
'codigo: código único identificador de la factura por la Aplicación externa.
'tipo: Tipo de documento enviado.
'email: Dirección del destinatario.
'asunto: Asunto del correo electrónico de notificación del correo.
'cuerpo: Cuerpo del correo electrónico (admite forma html). A este cuerpo se incluye automáticamente los datos de enlace hacia la plataforma de pago,
    'y el pie definido en la plataforma. En caso de no existir, no se envía el email.
'fecha: Fecha de la factura (formato yyyy-mm-dd).
'importe: Importe de la factura (formato nnnn.nn).


Set myAttachments = Item.Attachments ' Documentos adjuntos
Num_Adj = myAttachments.Count
If Num_Adj > 0 Then
    EnvioDoble = True
Else
    EnvioDoble = False
End If
    
asunto = Limpia_Comas(Item.Subject)
factura = ExtraeNumFactura(Item.Subject)
If Not ControlNumFactura(factura) Then
    CodError = "Numero de Documento"
    GoTo ErrorEnvia
End If
tipo = ""

EnviaUnoSeg = 0

LOPD1 = "<p><pre>" & "******************** ADVERTENCIA LEGAL ******************** "
LOPD2 = "<br>" & "Este mensaje contiene información confidencial destinada para ser leída exclusivamente por el destinatario."
LOPD3 = "<br>" & "Queda prohibida su reproducción, publicación y divulgación total o parcial del mensaje, así como el uso no autorizado por el emisor."
LOPD4 = "<br>" & "Si Vd. lo ha recibido por error, le rogamos que por favor lo destruya inmediatamente y se ponga en contacto con nosotros."
LOPD5 = "<br>" & "Su dirección de correo se encuentra recogida en nuestros ficheros con la finalidad de mantener correspondencia electrónica, "
LOPD6 = "<br>" & "responder a las consultas por Vd. planteadas y el envío de comunicaciones por diversos medios, incluyendo los electrónicos,"
LOPD7 = "<br>" & "entendiéndose que consiente el tratamiento de los citados datos con dicha finalidad.  Usted puede ejercitar sus derechos de acceso,"
LOPD8 = "<br>" & "rectificación, cancelación y oposición ante REGISTRO MERCANTIL DE SEVILLA CB de acuerdo a lo previsto en Reglamento General de Protección de Datos 2016/679"
LOPD9 = "<br>" & "del Parlamento Europeo y del Consejo, de 27 de abril de 2016 y en la Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales"
LOPD10 = "<br>" & "y garantía de los derechos digitales (BOE núm.  294, de 6 de diciembre de 2018)." & "</pre></p>"


If InStr(Item, "Despacho") <> 0 Then
            
    mail = ExtraeMail(Item.Recipients(1))

    marca = InStr(1, Item.body, "El pago se")
    If marca = 0 Then
        'MsgBox "Esta notificacion no tiene el formato esperado"
        'Exit Function
        If EnvioDoble Then
            tipo = "4"
            marca = Len(Item.body)
        End If
    End If
    body = Mid(Item.body, 1, marca - 1)
    body = Limpia_Espacios(Limpia_guiones(body))
               
    marca1_0 = InStr(body, ":") '
    marca1_1 = InStr(body, "/")
    marca2 = InStr(body, "PRESENTANTE:")
    marca3 = InStr(body, "SOCIEDAD:")
    marca4 = InStr(body, "PROTOCOLO:")
    marca5 = InStr(body, "AUTORIZANTE:")
    marca6 = InStr(body, "el cual ya ha sido")
    marca7 = InStr(body, "asciende a:") + 11
    marca8 = InStr(body, "Total Registro")
   
        
    If marca1_1 <> 0 And marca2 <> 0 Then
        bodyini = Mid(body, marca1_1 - 1, Len(body))
        bodyini = Mid(bodyini, 1, marca2 - marca1_1)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    
    If marca2 <> 0 And marca3 <> 0 Then
        bodypre = Mid(body, marca2, Len(body))
        bodypre = Mid(bodypre, 1, marca3 - marca2)
        'bodypre = TextoSinAcentos(bodypre)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    If marca3 <> 0 And marca4 <> 0 Then
        bodySoc = Mid(body, marca3, Len(body))
        bodySoc = Mid(bodySoc, 1, marca4 - marca3)
       ' bodySoc = TextoSinAcentos(bodySoc)
    ElseIf marca4 = 0 Then
        If marca5 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca5 - marca3)
            'bodySoc = TextoSinAcentos(bodySoc)
        ElseIf marca6 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca6 - marca3)
           ' bodySoc = TextoSinAcentos(bodySoc)
        End If
        'GoTo ErrorEnvia
    End If
    If marca4 <> 0 And marca5 <> 0 Then
        bodyProt = Mid(body, marca4, Len(body))
        bodyProt = Mid(bodyProt, 1, marca5 - marca4)
        bodyProt = bodyProt & "<br>"
    Else
        'GoTo ErrorEnvia
    End If
    If marca5 <> 0 And marca6 <> 0 Then
        bodyAut = Mid(body, marca5, Len(body))
        bodyAut = Mid(bodyAut, 1, marca6 - marca5)
        'bodyAut = TextoSinAcentos(bodyAut)
        bodyAut = bodyAut & "<br>"
    Else
        'GoTo ErrorEnvia
    End If
    If marca6 <> 0 And marca7 <> 0 Then
        bodyelcual = Mid(body, marca6, Len(body))
        bodyelcual = Mid(bodyelcual, 1, marca7 - marca6)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    
    If marca8 <> 0 Then
        bodyasciende = Mid(body, marca8, Len(body))
        marcae = InStr(bodyasciende, "euros")
        bodyasciende = Mid(bodyasciende, 1, marcae + 4)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    restodevolver = Mid(body, marca8, Len(body) - marca8)
    
    marcar1 = InStr(bodyasciende, "ro")
    bodyresto0 = Left(bodyasciende, marcar1 + 1)
    marcar2 = InStr(bodyasciende, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyasciende, 1, marcar2 + 4))
    bodyresto0 = bodyresto0 & ": " & bodyresto2 & " euros"
    
    marca9 = InStr(body, "Total Provisión") - 1
    bodyresto = Right(body, Len(body) - marca9)
    marcar1 = InStr(bodyresto, ":")
    bodyresto1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1 = bodyresto1 & " " & bodyresto2 & " euros"
    
    marca9_1 = InStr(body, "A PAGAR") - 1
    
    bodyresto = Right(body, Len(body) - marca9_1)
    marcar1 = InStr(bodyresto, "R")
    bodyresto1_1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
    'bodyunido = Mid(body, 1, marca1_0) & "<br>" & bodyini & "<br>" & bodypre & "</br>" & "<br>" & bodySoc & "</br>" & "<br>" & bodyProt & "</br>" & "<br>" & bodyAut & "</br>" & "<br>" & bodyElcual & "</br>" & "<br>" & bodyresto0 & "</br>" & "<br>" & bodyresto1 & "</br>" & "<br>" & bodyresto1_1 & "</br>"
    
    marca1 = InStr(body, "A PAGAR")
    If marca1 <> 0 Then
        Final_body = Mid(body, marca1, Len(body))
        body2 = "" '  "<br>Forma de Pago: Enlace abajo indicado."
        body2 = body2 & "<p>Para cualquier consulta o aclaración sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 954 54 20 93 ext 1," & "<br>" & "en la dirección de correo sevilla.gestion@registromercantil.org o en www.rmsevilla.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
                 
    Else
        marca1 = InStr(body, "A DEVOLVER")
        
        '****
        marca9_1 = InStr(body, "A DEVOLVER") - 1
    
        bodyresto = Right(body, Len(body) - marca9_1)
        marcar1 = InStr(bodyresto, "R")
        bodyresto1_1 = Left(bodyresto, marcar1)
        marcar2 = InStr(bodyresto, "euros")
        bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
        bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
        '********
        Final_body = Mid(body, marca1, Len(body))
        body2 = "" '"<br>Forma de devolución: Enlace abajo indicado."
        body2 = body2 & "<p>Para cualquier consulta o aclaración sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 954 54 20 93 ext 1," & "<br>" & "en la dirección de correo sevilla.gestion@registromercantil.org o en www.rmsevilla.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>" ' "<br>" & restodevolver & "<br>"
        'Replace(bodyelcual, ".", ".<br>")
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
    End If
                                                                
    importe = importe & ExtraeNumeros(Final_body)
                
    fecha = Format(Date, "dd/mm/yyyy")
    
    ControlWord factura, fecha, bodySoc, importe
    
    If tipo = "" Then
    Select Case AnalizaTipo(factura)
        Case 1: tipo = "1"
        Case 5: tipo = "5"
        Case 2: tipo = "2"
        Case Else
    End Select
    End If
    
    EnviaUnoSeg = EnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, "c:\GPFD\Aviso1.pdf")
    'EnviaUno = GpfdEnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, "c:\GPFD\Aviso1.pdf")
    If EnviaUno <> 0 Then
        If EnvioDoble Then
            LOPD1 = Chr(13) & "******************** ADVERTENCIA LEGAL ******************** " & Chr(13)
            LOPD2 = "Este mensaje contiene información confidencial destinada para ser leída exclusivamente por el destinatario." & Chr(13)
            LOPD3 = "Queda prohibida su reproducción, publicación y divulgación total o parcial del mensaje, así como el uso no autorizado por el emisor."
            LOPD4 = " Si Vd. lo ha recibido por error, le rogamos que por favor lo destruya inmediatamente y se ponga en contacto con nosotros." & Chr(13)
            LOPD5 = "Su dirección de correo se encuentra recogida en nuestros ficheros con la finalidad de mantener correspondencia electrónica,"
            LOPD6 = " responder a las consultas por Vd. planteadas y el envío de comunicaciones por diversos medios, incluyendo los electrónicos,"
            LOPD7 = " entendiéndose que consiente el tratamiento de los citados datos con dicha finalidad.  Usted puede ejercitar sus derechos de acceso,"
            LOPD8 = " rectificación, cancelación y oposición ante REGISTRO MERCANTIL DE SEVILLA CB de acuerdo a lo previsto en Reglamento General de Protección de Datos 2016/679"
            LOPD9 = " del Parlamento Europeo y del Consejo, de 27 de abril de 2016 y en la Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales"
            LOPD10 = " y garantía de los derechos digitales (BOE núm.  294, de 6 de diciembre de 2018)."
            
            body = Mid(Item.body, 1, marca - 1)
            marcaentrada = InStr(body, factura)
            
            marca1 = InStr(body, "A PAGAR")
            body = Mid(body, marcaentrada, marca)
            If marca1 <> 0 Then ' A PAGAR
                TextoAdicional = Chr(13) & "Le hemos enviado otro mail desde nuestra plataforma de cobros GPFD, con un enlace que le permitirá acceder al documento para realizar el pago por alguna de las diferentes pasarelas: tarjeta, banco, paypal." & Chr(13)
            Else
                TextoAdicional = Chr(13) & "Le hemos enviado otro mail desde nuestra plataforma de pagos GPFD, con un enlace que le permitirá acceder al documento e indicarnos sus datos bancarios para realizarle dicha devolución." & Chr(13)
            End If
            '!!!!!!!!!!!!!!!!!!!!!!!!no vale (coger del de arriba sin los br
            nuevobody = "Le informamos que se encuentra despachado el documento telemático con número de entrada: " & body & TextoAdicional & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
            Item.body = nuevobody
            Item.Send
        Else
            Item.Move (olFolder.Folders("GPFD"))
        End If
    End If
    
ElseIf InStr(Item, "Depósito") <> 0 Then
    
    mail = ExtraeMail(Item.Recipients(1))

    marca = InStr(1, Item.body, "El pago se")
    If marca = 0 Then
        MsgBox "Esta notificacion no tiene el formato esperado"
        Exit Function
    End If
    body = Mid(Item.body, 1, marca - 1)
    body = Limpia_Espacios(Limpia_guiones(body))
               
    marca1_0 = InStr(body, ":") '
    marca1_1 = InStr(body, "/")
    marca1_2 = InStr(body, ",0")
    marca2 = InStr(body, "PRESENTANTE:")
    marca3 = InStr(body, "SOCIEDAD:")
    marca4 = InStr(body, "El importe")
    marca7 = InStr(body, "asciende a:") + 11
    marca8 = InStr(body, "Total Registro")
   
        
    If marca1_1 <> 0 And marca1_2 <> 0 Then
        bodyini = Mid(body, marca1_1 - 1, Len(body))
        bodyini = Mid(bodyini, 1, marca1_2 + 3 - marca1_1)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    
    If marca2 <> 0 And marca3 <> 0 Then
        bodypre = Mid(body, marca2, Len(body))
        bodypre = Mid(bodypre, 1, marca3 - marca2)
        'bodypre = TextoSinAcentos(bodypre)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    If marca3 <> 0 And marca4 <> 0 Then
        bodySoc = Mid(body, marca3, Len(body))
        bodySoc = Mid(bodySoc, 1, marca4 - marca3)
       ' bodySoc = TextoSinAcentos(bodySoc)
    ElseIf marca4 = 0 Then
        If marca5 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca5 - marca3)
            'bodySoc = TextoSinAcentos(bodySoc)
        ElseIf marca6 <> 0 Then
            bodySoc = Mid(body, marca3, Len(body))
            bodySoc = Mid(bodySoc, 1, marca6 - marca3)
           ' bodySoc = TextoSinAcentos(bodySoc)
        End If
        'GoTo ErrorEnvia
    End If
    
    If marca8 <> 0 Then
        bodyasciende = Mid(body, marca8, Len(body))
        marcae = InStr(bodyasciende, "euros")
        bodyasciende = Mid(bodyasciende, 1, marcae + 4)
    Else
        CodError = "Faltan datos."
        GoTo ErrorEnvia
    End If
    restodevolver = Mid(body, marca8, Len(body) - marca8)
    
    marcar1 = InStr(bodyasciende, "ro")
    bodyresto0 = Left(bodyasciende, marcar1 + 1)
    marcar2 = InStr(bodyasciende, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyasciende, 1, marcar2 + 4))
    bodyresto0 = bodyresto0 & ": " & bodyresto2 & " euros"
    
    marca9 = InStr(body, "Total Provisión") - 1
    bodyresto = Right(body, Len(body) - marca9)
    marcar1 = InStr(bodyresto, ":")
    bodyresto1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1 = bodyresto1 & " " & bodyresto2 & " euros"
    
    marca9_1 = InStr(body, "A PAGAR") - 1
    
    bodyresto = Right(body, Len(body) - marca9_1)
    marcar1 = InStr(bodyresto, "R")
    bodyresto1_1 = Left(bodyresto, marcar1)
    marcar2 = InStr(bodyresto, "euros")
    bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
    bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
    'bodyunido = Mid(body, 1, marca1_0) & "<br>" & bodyini & "<br>" & bodypre & "</br>" & "<br>" & bodySoc & "</br>" & "<br>" & bodyProt & "</br>" & "<br>" & bodyAut & "</br>" & "<br>" & bodyElcual & "</br>" & "<br>" & bodyresto0 & "</br>" & "<br>" & bodyresto1 & "</br>" & "<br>" & bodyresto1_1 & "</br>"
    
    marca1 = InStr(body, "A PAGAR")
    If marca1 <> 0 Then
        Final_body = Mid(body, marca1, Len(body))
        body2 = "" '  "<br>Forma de Pago: Enlace abajo indicado."
        body2 = body2 & "<p>Para cualquier consulta o aclaración sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 954 54 20 93 ext 1," & "<br>" & "en la dirección de correo sevilla.gestion@registromercantil.org o en www.rmsevilla.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>"
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
                 
    Else
        marca1 = InStr(body, "A DEVOLVER")
        
        '****
        marca9_1 = InStr(body, "A DEVOLVER") - 1
    
        bodyresto = Right(body, Len(body) - marca9_1)
        marcar1 = InStr(bodyresto, "R")
        bodyresto1_1 = Left(bodyresto, marcar1)
        marcar2 = InStr(bodyresto, "euros")
        bodyresto2 = ExtraeNumeros(Mid(bodyresto, 1, marcar2 + 4))
        bodyresto1_1 = bodyresto1_1 & ": " & bodyresto2 & " euros"
        '********
        Final_body = Mid(body, marca1, Len(body))
        body2 = "" '"<br>Forma de devolución: Enlace abajo indicado."
        body2 = body2 & "<p>Para cualquier consulta o aclaración sobre el documento despachado puede ponerse en contacto con nosotros a través del teléfono 954 54 20 93 ext 1," & "<br>" & "en la dirección de correo sevilla.gestion@registromercantil.org o en www.rmsevilla.com.</p>"
        bodyunido = Mid(body, 1, marca1_0) & bodyini & "<br>" & bodypre & "<br>" & bodySoc & "<br>" & bodyProt & bodyAut & Replace(bodyelcual, ".", ".<br>") & "<br>" & bodyresto0 & " " & bodyresto1 & " " & bodyresto1_1 & "<br>" ' "<br>" & restodevolver & "<br>"
        'Replace(bodyelcual, ".", ".<br>")
        body1 = bodyunido & body2
        body1 = body1 & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
    End If
                                                                
    importe = importe & ExtraeNumeros(Final_body)
                
    fecha = Format(Date, "dd/mm/yyyy")
    
    ControlWord factura, fecha, bodySoc, importe
    
    If tipo = "" Then
    Select Case AnalizaTipo(factura)
        Case 1: tipo = "1"
        Case 5: tipo = "5"
        Case 2: tipo = "2"
        Case Else
    End Select
    End If
    
    EnviaUnoSeg = EnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, "c:\GPFD\Aviso1.pdf")
    If EnviaUno <> 0 Then
        If EnvioDoble Then
            LOPD1 = Chr(13) & "******************** ADVERTENCIA LEGAL ******************** " & Chr(13)
            LOPD2 = "Este mensaje contiene información confidencial destinada para ser leída exclusivamente por el destinatario." & Chr(13)
            LOPD3 = "Queda prohibida su reproducción, publicación y divulgación total o parcial del mensaje, así como el uso no autorizado por el emisor."
            LOPD4 = " Si Vd. lo ha recibido por error, le rogamos que por favor lo destruya inmediatamente y se ponga en contacto con nosotros." & Chr(13)
            LOPD5 = "Su dirección de correo se encuentra recogida en nuestros ficheros con la finalidad de mantener correspondencia electrónica,"
            LOPD6 = " responder a las consultas por Vd. planteadas y el envío de comunicaciones por diversos medios, incluyendo los electrónicos,"
            LOPD7 = " entendiéndose que consiente el tratamiento de los citados datos con dicha finalidad.  Usted puede ejercitar sus derechos de acceso,"
            LOPD8 = " rectificación, cancelación y oposición ante REGISTRO MERCANTIL DE SEVILLA CB de acuerdo a lo previsto en Reglamento General de Protección de Datos 2016/679"
            LOPD9 = " del Parlamento Europeo y del Consejo, de 27 de abril de 2016 y en la Ley Orgánica 3/2018, de 5 de diciembre, de Protección de Datos Personales"
            LOPD10 = " y garantía de los derechos digitales (BOE núm.  294, de 6 de diciembre de 2018)."

            body = Mid(Item.body, 1, marca - 1)
            marca1 = InStr(body, "A PAGAR")
            If marca1 <> 0 Then ' A PAGAR
                TextoAdicional = Chr(13) & "Le hemos enviado otro mail desde nuestra plataforma de cobros GPFD, con un enlace que le permitirá acceder al documento para realizar el pago por alguna de las diferentes pasarelas: tarjeta, banco, paypal." & Chr(13)
            Else
                TextoAdicional = Chr(13) & "Le hemos enviado otro mail desde nuestra plataforma de pagos GPFD, con un enlace que le permitirá acceder al documento e indicarnos sus datos bancarios para realizarle dicha devolución." & Chr(13)
            End If
            nuevobody = body & TextoAdicional & LOPD1 & LOPD2 & LOPD3 & LOPD4 & LOPD5 & LOPD6 & LOPD7 & LOPD8 & LOPD9 & LOPD10
            Item.body = nuevobody
            Item.Send
        End If
        Item.Move (olFolder.Folders("GPFD"))
    Else
        GoTo ErrorEnvia
    End If
    
'ElseIf InStr(Item, "Factura") <> 0 Then
'    mail = ExtraeMail(Item.Recipients(1))
'    fecha = Format(Date, "dd/mm/yyyy")
'    tipo = "7"
'    body1 = Item.body
'    'importe = importe & ExtraeNumeros(body1)
'    miAdjunto = saveAttachtoDisk(Item)
'    EnviaUno = EnvioFacturas(apikey, factura, tipo, mail, asunto, body1, fecha, importe, miAdjunto)

End If

Exit Function

ErrorEnvia:
EnviaUnoSeg = 0
MsgBox "No se ha podido enviar la notificacion. " & factura & "  Error:" & CodError

End Function
Sub SaveAttachment()
 
Dim myInspector As Outlook.Inspector
Dim myItem As Outlook.MailItem
Dim myAttachments As Outlook.Attachments
  
 
 Set myInspector = Application.ActiveInspector
 If Not TypeName(myInspector) = "Nothing" Then
 
    If TypeName(myInspector.CurrentItem) = "MailItem" Then
    
        Set myItem = myInspector.CurrentItem
        
        Set myAttachments = myItem.Attachments
        
        'Prompt the user for confirmation
        
        Dim strPrompt As String
        
        strPrompt = "Are you sure you want to save the first attachment in the current item to the Documents folder? If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file."
        
        If MsgBox(strPrompt, vbYesNo + vbQuestion) = vbYes Then
        
           myAttachments.Item(1).SaveAsFile Environ("HOMEPATH") & "\My Documents\" & myAttachments.Item(1).DisplayName
        
        End If
        
    Else
        
        MsgBox "The item is of the wrong type."
    
    End If
 
 End If
 
End Sub

Function AnalizaTipo(factura As String) As String
    AnalizaTipo = Left(factura, 1)
End Function

Function ExtraeNumFactura(texto As String) As String
Dim par1, par2 As Long
Dim temp As String

par1 = InStr(texto, "(")
par2 = InStr(texto, ")")
If par1 = 0 Or par2 = 0 Then
    ExtraeNumFactura = ""
Else
    temp = Mid(texto, par1 + 1, par2 - par1 - 1)
    ExtraeNumFactura = Limpia_EspaciosUno(temp)
End If


End Function

Function ControlNumFactura(texto As String) As Boolean
Dim b1, b2, c1 As Long
Dim temp1, temp2, temp3, temp4 As String

If texto <> "" And Not IsNull(texto) Then
    b1 = InStr(texto, "/")
    If b1 > 0 Then ' hay una barra
        temp1 = Mid(texto, 1, b1 - 1) ' cojo la primera parte
        If IsNumeric(temp1) Then ' si es un numero sigo
            temp2 = Mid(texto, b1 + 1, Len(texto) - b1) ' cojo las dos siguientes partes
            b2 = InStr(temp2, "/")
            If b2 > 0 Then ' hay otra barra
                temp3 = Mid(temp2, 1, b2 - 1) 'cojo la parte central
                If IsNumeric(temp3) Then ' si es un numero sigo
                    temp4 = Mid(temp2, b2 + 1, Len(temp2) - b2) ' obtengo la parte final
                    c1 = InStr(temp4, ",")
                    If c1 > 0 Then ' hay coma en la parte final
                        temp5 = Mid(temp4, c1 + 1, Len(temp4) - c1) ' obtengo lo que hay detras de la coma
                        If temp5 = "0" Then ControlNumFactura = True
                    End If
                End If
            End If
        End If
    End If
    
Else
    ControlNumFactura = False
End If

End Function


Public Function EnvioFacturas(apikey As String, factura As String, tipo As String, email As String, asunto As String, body As String, fecha As Date, importe As String, arch As String) As Long

Dim JSON As String
Dim cadena As String
Dim NumArchivo2 As Integer
Dim vnombreFichero2 As String
Dim apik As String
Dim NF As String
Dim em As String
Dim asun As String
Dim fec As String
Dim Fecha_E As String
Dim impor As String
Dim ImporP As String
Dim fac As String
Dim bod As String
Dim tip As String
Dim id_devuelto As String
Dim http As String

'pruebas rafa
'http://gpfd.rafatux.com/
'http = "http://gpfd.test.rafatux.com/api/documento_nuevo.php"

'sandbox
'http = "http://gpfd.sitelcom.es/api/documento_nuevo.php"

'produccion actual
http = "https://www.gpfd.es/api/documento_nuevo.php" '2021

ImporP = importe
If TieneComa(importe) <> 0 Then
    If Len(DigitosComa(importe)) = 1 Then ImporP = importe & "0"
Else
    ImporP = importe & ".00"
End If

Fecha_E = Format(fecha, "DD-MM-YYYY")

body = SustituyeTildes(body)
JSON = Upload(http, arch, "archivo", "apikey=" & apikey & "|codigo=" & factura & "|tipo=" & tipo & "|destinatario=" & email & "|asunto=" & asunto & "|fecha_documento=" & Fecha_E & "|cuerpo=" & body & "|importe=" & ImporP)

If InStr(1, JSON, "id") <> 0 Then 'if Not IsEmpty(JSON) And JSON <> "" And Not IsNull(JSON) Then

    'If Left(JSON, 1) = "{" And Right(JSON, 1) = "}" Then
        id_devuelto = Devuelve_id_Subida(JSON)
        If id_devuelto = "0" Then
             MsgBox "NO se ha podido realizar el envío." & Chr(13) & "{" & Right(JSON, Len(JSON) - InStr(1, JSON, "error:") - 10), vbCritical
        End If
    'End If
         
    EnvioFacturas = ExtraeNumeros(id_devuelto)
    
Else

    'JSON vacio
    MsgBox "NO se ha podido realizar el envío.", vbCritical
    EnvioFacturas = 0
    
End If


End Function

Function Devuelve_id_Subida(t As String) As String

    Dim S1 As String
    Dim textstart As Integer
    Dim textend As Integer
    
    t = Limpia_EspaciosUno(t)
    textstart = InStr(1, t, "error") - 2
    textend = InStr(1, t, ":") + 1
    longitud = textstart - textend
    S1 = Mid(t, textend, longitud)
    'S1 = Right(S1, Len(S1) - 9)
   
    Devuelve_id_Subida = S1

End Function

Function TieneComa(nombre As String) As String

TieneComa = InStr(1, nombre, ",")

If TieneComa = "0" Then
    TieneComa = InStr(1, nombre, ".")
End If

End Function

Function DigitosComa(nombre As String) As String

Dim x As String

x = TieneComa(nombre)
If x <> 0 Then
    'DigitosComa = Left(Nombre, X - 1)
    DigitosComa = Right(nombre, Len(nombre) - x)
Else
    DigitosComa = "00"
End If

End Function

Function wshRun(Command As String, Optional WindowStyle, Optional WaitOnReturn) As Long

Dim wShell As Object 'wshShell
Dim fso As Object 'FileSystemObject

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Command = fso.GetFile(Command).ShortPath
    Set fso = Nothing
    
    On Error GoTo err_wshRun
    Set wShell = CreateObject("WScript.Shell")
    wshRun = wShell.Run(Command, WindowStyle, WaitOnReturn)
    Set wShell = Nothing
    wshRun = 0

    Exit Function
    
err_wshRun:
    
    MsgBox Command & ". Archivo o ruta no encontrado.", vbCritical '& " " & err.Description
    wshRun = 1
End Function

Public Function TextoSinAcentos(ByVal texto As String) As String
Dim strNormalizado As String
'strNormalizado = Replace(texto, "Á", " ")
strNormalizado = Replace(texto, "Á", "&#193;")
strNormalizado = Replace(strNormalizado, "É", "&#201;")
strNormalizado = Replace(strNormalizado, "Í", "&#205;")
strNormalizado = Replace(strNormalizado, "Ó", "&#211;")
strNormalizado = Replace(strNormalizado, "Ú", "&#218;")
strNormalizado = Replace(strNormalizado, "á", "&#225;")
strNormalizado = Replace(strNormalizado, "é", "&#233;")
strNormalizado = Replace(strNormalizado, "í", "&#237;")
strNormalizado = Replace(strNormalizado, "ó", "&#243;")
strNormalizado = Replace(strNormalizado, "ú", "&#250;")
strNormalizado = Replace(strNormalizado, "ñ", "&ntilde;")
strNormalizado = Replace(strNormalizado, "Ñ", "&Ntilde;")

TextoSinAcentos = strNormalizado
End Function

Public Function SustituyeTildes(ByVal texto As String) As String
' Esta función devuelve el texto sin acentos

Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado As String

lngTexto = Len(texto)
If lngTexto = 0 Then
    SustituyeTildes = ""
    Exit Function
End If

strNormalizado = Replace(texto, "Á", "&Aacute;")
strNormalizado = Replace(strNormalizado, "á", "&aacute;")
strNormalizado = Replace(strNormalizado, "É", "&Eacute;")
strNormalizado = Replace(strNormalizado, "é", "&eacute;")
strNormalizado = Replace(strNormalizado, "Í", "&Iacute;")
strNormalizado = Replace(strNormalizado, "í", "&iacute;")
strNormalizado = Replace(strNormalizado, "Ó", "&Oacute;")
strNormalizado = Replace(strNormalizado, "ó", "&oacute;")
strNormalizado = Replace(strNormalizado, "Ú", "&Uacute;")
strNormalizado = Replace(strNormalizado, "ú", "&uacute;")
strNormalizado = Replace(strNormalizado, "Ý", "&Yacute;")
strNormalizado = Replace(strNormalizado, "ý", "&yacute;")
strNormalizado = Replace(strNormalizado, "ñ", "&ntilde;")
strNormalizado = Replace(strNormalizado, "Ñ", "&Ntilde;")


SustituyeTildes = strNormalizado

End Function

Public Function Limpia_Espacios(ByVal texto As String) As String

Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado, strNormalizado1, strNormalizado2 As String

lngTexto = Len(texto)
If lngTexto = 0 Then
    Limpia_Espacios = ""
    Exit Function
End If

strNormalizado = texto
For i = 1 To 10
    strNormalizado = Replace(strNormalizado, "  ", " ")  ' Elimina
Next i

strNormalizado1 = Replace(strNormalizado, vbTab, "")  ' Elimina
strNormalizado2 = Replace(strNormalizado1, vbLf, "")  ' Elimina
strNormalizado2 = Replace(strNormalizado2, vbCr, "")
strNormalizado2 = Replace(strNormalizado2, vbNewLine, "")
strNormalizado2 = Replace(strNormalizado2, vbNullString, "")
strNormalizado2 = Replace(strNormalizado2, vbNullChar, "")


Limpia_Espacios = strNormalizado2
End Function

Public Function Limpia_guiones(ByVal texto As String) As String

Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado As String

lngTexto = Len(texto)
If lngTexto = 0 Then
    Limpia_guiones = ""
    Exit Function
End If

strNormalizado = texto
For i = 1 To 10
    strNormalizado = Replace(strNormalizado, "--", "-")  ' Elimina
Next i
Limpia_guiones = strNormalizado

End Function

Public Function Limpia_EspaciosUno(ByVal texto As String) As String

Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado As String

lngTexto = Len(texto)
If lngTexto = 0 Then
    Limpia_EspaciosUno = ""
    Exit Function
End If

strNormalizado = texto
For i = 1 To 3
    strNormalizado = Replace(strNormalizado, " ", "") ' Elimina enter
Next i
Limpia_EspaciosUno = strNormalizado
End Function

Public Function Cambia_Espacios_por_Guion(ByVal texto As String) As String

Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado As String

lngTexto = Len(texto)
If lngTexto = 0 Then
    Cambia_Espacios_por_Guion = ""
    Exit Function
End If

strNormalizado = texto
For i = 1 To 10
    strNormalizado = Replace(texto, " ", "-") ' Elimina enter
Next i
Cambia_Espacios_por_Guion = strNormalizado
End Function

Private Function ExtraeNumeros(cadena As String) As String
'busco la cadena: importe a devolver
'si la encuentro es que la cantidad es a devolver y el signo es negativo
Dim Signo As String, n As Long, c As String, r As String
Dim flag As Boolean
flag = False
Signo = ""
If InStr(1, cadena, "a devolver", vbTextCompare) <> 0 Then Signo = "-"
   
   For n = 1 To Len(cadena)
      c = Mid(cadena, n, 1)
      Select Case c
      Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
         r = r & c
         flag = True
      Case ","
        If flag Then
          r = r & "."
          flag = False
        End If
      Case "."
        If flag Then
          'r = r & "."
          flag = False
        End If
      End Select
   Next
   
   ExtraeNumeros = Signo & r
End Function

Private Function ExtraeMail(cadena As String) As String
   Dim n As Long, c As String, r As String
   For n = 1 To Len(cadena)
      c = Mid(cadena, n, 1)
      Select Case c
      Case Is <> "'"
         r = r & c
      End Select
   Next
   ExtraeMail = r
End Function

Private Function Limpia_Comas(cadena As String) As String
   Dim n As Long, c As String, r As String
   
   For n = 1 To Len(cadena)
      c = Mid(cadena, n, 1)
      Select Case c
      Case Is <> "'"
         r = r & c
      Case Is = ","
        r = r & "."
      End Select
   Next
   Limpia_Comas = r
End Function

Private Function Signo(cadena As String) As String
'busco la cadena: importe a devolver
' si la encuentro es que la cantidad es a devolver y el signo es negativo
Signo = ""
If InStr(1, cadena, "importe a devolver", vbTextCompare) <> 0 Then Signo = "-"

End Function

Sub ControlWord(factura As String, fecha As Date, nombre As String, importe As String)
'Dim MyData As Object
    
    Dim wdRange1 As Word.Range
    Dim wdRange2 As Word.Range
    Dim wdRange3 As Word.Range
    Dim wdRange4 As Word.Range
       
    Dim appWD As Word.Application
    ' Create a new instance of Word and make it visible
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = False
   
        ' Tell Word to create a new document
        appWD.Documents.Open "c:\gpfd\aviso.docx"
        
        Set wdRange1 = appWD.ActiveDocument.Bookmarks("Texto1").Range
        wdRange1.Text = factura
        Set wdRange2 = appWD.ActiveDocument.Bookmarks("Texto2").Range
        wdRange2.Text = nombre
        Set wdRange3 = appWD.ActiveDocument.Bookmarks("Texto3").Range
        wdRange3.Text = fecha
        Set wdRange4 = appWD.ActiveDocument.Bookmarks("Texto4").Range
        wdRange4.Text = importe & "€"
        
        ' Save the new document
        appWD.ActiveDocument.ExportAsFixedFormat "c:\gpfd\Aviso1.pdf", wdExportFormatPDF
        ' Close the new Word document.
        appWD.ActiveDocument.Close False
 
    ' Close the Word application.
    appWD.Quit
End Sub

Private Sub Application_Startup()
  Envia
End Sub


Function Upload(strUploadUrl, strFilePath, strFileField, strDataPairs)
'Uses POST to upload a file and miscellaneous form data
'strUploadUrl is the URL (http://127.0.0.1/cgi-bin/upload.exe)
'strFilePath is the file to upload (C:\My Documents\test.zip)
'strFileField is the web page equivalent form field name for the file (File1)
'strDataPairs are pipe-delimited form data pairs (foo=bar|snap=crackle)
On Error GoTo Err_upload

Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
Dim nFile           As Integer
Dim baBuffer()      As Byte
Dim sPostData       As String
Dim sPostData1      As String
Dim strFormStart, strDataPair
Dim respuesta
Dim web
'Create the multipart form data
    
    respuesta = ""
    'First add any ordinary form data pairs
    strFormStart = ""
    For Each strDataPair In Split(strDataPairs, "|")
        strFormStart = strFormStart & "--" & STR_BOUNDARY & vbCrLf
        strFormStart = strFormStart & "Content-Disposition: form-data; "
        strFormStart = strFormStart & "name=""" & Split(strDataPair, "=")(0) & """"
        strFormStart = strFormStart & vbCrLf & vbCrLf
        strFormStart = strFormStart & Split(strDataPair, "=")(1)
        strFormStart = strFormStart & vbCrLf
    Next
    'Now add the header for the uploaded file
    strFormStart = strFormStart & "--" & STR_BOUNDARY & vbCrLf
    strFormStart = strFormStart & "Content-Disposition: form-data; "
    strFormStart = strFormStart & "name=""" & strFileField & """; "
    strFormStart = strFormStart & "filename=""" & Mid(strFilePath, InStrRev(strFilePath, "\") + 1) & """"
    strFormStart = strFormStart & vbCrLf
    strFormStart = strFormStart & "Content-Type: application/pdf" 'upload" 'bogus, but it works
    strFormStart = strFormStart & vbCrLf & vbCrLf
     
        '--- read file
    nFile = FreeFile
    Open strFilePath For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData1 = StrConv(baBuffer, vbUnicode) ' aqui esta el fichero
    End If
    Close nFile
    
    '--- prepare body
    sPostData = strFormStart & sPostData1 & vbCrLf & "--" & STR_BOUNDARY & "--" & vbCrLf
 
    Set web = CreateObject("Microsoft.XMLHTTP")

    web.Open "POST", strUploadUrl, False
    web.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
    web.Send pvToByteArray(sPostData)
    
    Upload = web.responseText
    
Exit Function
Err_upload:
MsgBox err.Description
End Function

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Public Sub Espera(segundos As Integer)
    Dim Wait
    Wait = DateAdd("s", segundos, Now)
    While Now < Wait
    Wend
End Sub

Public Function saveAttachtoDisk(itm As Outlook.MailItem) As String
Dim objAtt As Outlook.Attachment

     For Each objAtt In itm.Attachments
          objAtt.SaveAsFile "C:\GPFD\" & objAtt.DisplayName
          saveAttachtoDisk = "C:\GPFD\" & objAtt.DisplayName
          Set objAtt = Nothing
     Next
End Function

