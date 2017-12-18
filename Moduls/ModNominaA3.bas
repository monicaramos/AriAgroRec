Attribute VB_Name = "ModNominaA3"
Option Explicit


Public Function GeneraFicheroA3(Contador As Long, FechaPago As Date) As Boolean
Dim Regs As Integer
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim RegImpBruto As String
Dim RegImpSS As String
Dim RegImpIRPF As String

Dim Importe As String
Dim FecPag As String

'Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraFicheroA3 = False

    NFic = -1

    NFic = FreeFile
    Open App.Path & "\anticipoA3.txt" For Output As NFic

    cad = "03" & "00017" & "00000" ' tipo de registro + codigo de empresa + centro o codigo de trabajador
    
    FecPag = Format(Year(FechaPago), "0000") & Format(Month(FechaPago), "00") & Format(Day(FechaPago), "00")

    SQL = "select * from rrecibosnomina where fechahora = " & DBSet(FechaPago, "F") & " and idcontador = " & DBSet(Contador, "N")
    ' añado la condicion de q solo se pasa a A3 si no hay embargo
    SQL = SQL & " and hayembargo = 0 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Regs = 0
    While Not Rs.EOF
    
        Regs = Regs + 1
        
        ' importe bruto
        Importe = Format(Int(DBLet(Rs!Importe, "N")), "00000") & Format((DBLet(Rs!Importe, "N") - Int(DBLet(Rs!Importe, "N"))) * 100, "00")
        If DBLet(Rs!Importe, "N") >= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpBruto = cad & Format(Rs!CodTraba, "000000") & FecPag & "001" & "250" & Importe 'cad+codtraba+fecha+incidencia+250+importe bruto
        Print #NFic, RegImpBruto
        
        ' dto seguridad social
        Importe = Format(Int(DBLet(Rs!importesegso1, "N")), "00000") & Format((DBLet(Rs!importesegso1, "N") - Int(DBLet(Rs!importesegso1, "N"))) * 100, "00")
        If DBLet(Rs!importesegso1, "N") <= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpSS = cad & Format(Rs!CodTraba, "000000") & FecPag & "001" & "255" & Importe 'cad+codtraba+fecha+incidencia+255+dtoss
        Print #NFic, RegImpSS
        
        
        ' irpf
        Importe = Format(Int(DBLet(Rs!importeirpf, "N")), "00000") & Format((DBLet(Rs!importeirpf, "N") - Int(DBLet(Rs!importeirpf, "N"))) * 100, "00")
        If DBLet(Rs!importeirpf, "N") <= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpIRPF = cad & Format(Rs!CodTraba, "000000") & FecPag & "001" & "256" & Importe 'cad+codtraba+fecha+incidencia+256+irpf
        Print #NFic, RegImpIRPF
    
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    

    
    Close (NFic)
    NFic = -1
    
    If Regs > 0 Then GeneraFicheroA3 = True
    Exit Function
    
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function



Public Function GeneraNominaA3(FechaPago As Date) As Boolean
Dim Regs As Integer
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim RegImpBruto As String
Dim RegDias As String

Dim Importe As String
Dim Dias As Integer
Dim FecPag As String

'Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraNominaA3 = False

    NFic = -1

    NFic = FreeFile
    Open App.Path & "\nominaA3.txt" For Output As NFic

    cad = "03" & "00017" & "00000" ' tipo de registro + codigo de empresa + centro o codigo de trabajador
    
    FecPag = Format(Year(FechaPago), "0000") & Format(Month(FechaPago), "00") & Format(Day(FechaPago), "00")

    SQL = "select * from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Regs = 0
    While Not Rs.EOF
        ' para cada trabajador he de generar 2 registros (
    
        Regs = Regs + 1
        
        ' importe bruto
        Importe = Format(Int(DBLet(Rs!importe1, "N")), "00000") & Format((DBLet(Rs!importe1, "N") - Int(DBLet(Rs!importe1, "N"))) * 100, "00")
        If DBLet(Rs!importe1, "N") >= 0 Then
            Importe = Importe & "+"
        Else
            Importe = Importe & "-"
        End If
        
        Importe = Importe & "000000000+"
        
        RegImpBruto = cad & Format(Rs!Codigo1, "000000") & FecPag & "001" & "001" & Importe 'cad+codtraba+fecha+incidencia+001+importe bruto
        Print #NFic, RegImpBruto
        
        ' dias trabajados
        Dias = Format(Int(DBLet(Rs!importe2, "N")), "00")
        
        RegDias = cad & Format(Rs!Codigo1, "000000") & FecPag & "016" & Format(Dias, "00") & "00" & Left(DBLet(Rs!Nombre1, "T") & "NNN", 31) & "00000000000000" 'cad+codtraba+fecha+016+dias+00+SSNNS..+"
        Print #NFic, RegDias
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Close (NFic)
    NFic = -1
    
    If Regs > 0 Then GeneraNominaA3 = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function




Public Function CopiarFicheroA3(vFicher As String, vFecha As String) As Boolean
' vFicher viene nombre.txt
Dim nomFich As String
Dim FicherAux As String


On Error GoTo ecopiarfichero

    CopiarFicheroA3 = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

'    Me.CommonDialog1.DefaultExt = "txt"
'
'    CommonDialog1.Filter = "Archivos txt|txt|"
'    CommonDialog1.FilterIndex = 1
'
'    ' copiamos el primer fichero
'    CommonDialog1.FileName = "anticipoA3.txt"
'    Me.CommonDialog1.ShowSave
    
    If Dir("c:\ariadna\enlaceA3", vbDirectory) <> "" Then
    
        FicherAux = Replace(vFicher, ".txt", "") & Format(vFecha, "yyyymmdd")
        
        Dim i As Integer
        Dim B As Boolean
        Dim FicherAux1 As String
        
        i = 0
        B = True
        While Dir("C:\Ariadna\EnlaceA3\" & FicherAux & ".txt", vbArchive) <> "" And B
            i = i + 1
            FicherAux1 = Replace(vFicher, ".txt", "") & Format(vFecha, "yyyymmdd") & "_" & i
            If Dir("C:\Ariadna\EnlaceA3\" & FicherAux1 & ".txt", vbArchive) = "" Then B = False
            FicherAux = FicherAux1
        Wend
        FileCopy App.Path & "\" & vFicher, "C:\Ariadna\EnlaceA3\" & FicherAux & ".txt"
        
        'FileCopy App.Path & "\" & vFicher, "C:\Ariadna\EnlaceA3\" & Replace(vFicher, ".txt", "") & Format(vFecha, "yyyymmdd") & ".txt"

    Else
        MkDir ("c:\ariadna\enlaceA3")
        FileCopy App.Path & "\" & vFicher, "C:\Ariadna\EnlaceA3\" & Replace(vFicher, ".txt", "") & Format(vFecha, "yyyymmdd") & ".txt"
    End If
    
    CopiarFicheroA3 = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function

