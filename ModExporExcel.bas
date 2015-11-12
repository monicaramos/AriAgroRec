Attribute VB_Name = "ModExporExcel"
'Option Explicit
'
'
'' ------------------------------------------------------------------
'' \\ -- Función para exportar los datos
'' ------------------------------------------------------------------
'
'
'Function Exportar_Excel( _
'    sFileName As String, _
'    ListView As ListView, _
'    Optional Progressbar As Progressbar, _
'    Optional SheetIndex As Integer = 1) As Boolean
'
'    On Error GoTo error_Handler
'
'    Dim obj_Excel       As Object               ' -- CREAR EL OBJETO (INSTANCIAR)CON EL OBJETO APLICACION (obj_Excel)
'    Dim obj_Libro       As Object
'
'    Dim iCol            As Integer              ' -- Variables para las columnas y filas
'    Dim iRow            As Integer
'
'    ' --  Nueva referencia a Excel y nuevo referencia al Libro
'    Set obj_Excel = CreateObject("Excel.Application")
'    With obj_Excel
'        ' -- Abrir el libro
'        Set obj_Libro = .Workbooks.Open(sFileName)
'    End With
'
'    With obj_Libro
'
'        ' -- Asignamos El valor Maximo del Progress teniendo
'        ' -- como dato la cantidad de items en el ListView
'        If Not Progressbar Is Nothing Then
'            Progressbar.Max = ListView.ListItems.Count
'            If Not Progressbar.visible Then Progressbar.visible = True
'        End If
'
'        ' -- Referencia a la hoja con índice 1
'        With .Sheets(SheetIndex)
'            ' -- Recorremos la cantidad de items del ListView
'            For iRow = 1 To ListView.ListItems.Count
'                iCol = 1
'                ' -- Asignamos el item actual en la celda
'                .Cells(iRow, iCol) = ListView.ListItems.item(iRow)
'
'                ' -- Asignamos el subitem actual en la celda
'                 For iCol = 1 To ListView.ColumnHeaders.Count - 1
'                     .Cells(iRow, iCol + 1) = ListView.ListItems(iRow).SubItems(iCol)
'                 Next
'
'                 If Not Progressbar Is Nothing Then
'                     ' -- Aumentamos en 1 la propiedad value
'                     Progressbar.Value = Progressbar.Value + 1
'                 End If
'            Next
'        End With
'    End With
'
'    ' -- Destruimos las variables de objeto
'    obj_Excel.visible = True
'    Set obj_Libro = Nothing
'    Set obj_Excel = Nothing
'    ' -- Ok
'    Exportar_Excel = True
'
'    If Not Progressbar Is Nothing Then
'       Progressbar.Value = 0
'       Progressbar.visible = False
'    End If
'
'
'' -- Errores
'Exit Function
'error_Handler:
'
'Exportar_Excel = False
'MsgBox Err.Description, vbCritical
'
'On Error Resume Next
'Set obj_Libro = Nothing
'Set obj_Excel = Nothing
'
'Progressbar.Value = 0
'End Function
'
'
'' ------------------------------------------------------------------
'' \\ -- Botón para exportar los datos al libro
'' ------------------------------------------------------------------
'Private Sub Command1_Click()
'
'    Dim ret As Boolean
'
'    ' -- Le pasa el path donde está ubicado el libro, _
'      -- el control ListView, opcional un Progressbar, _
'      -- y lo exporta en la hoja con índice 2
'    ret = Exportar_Excel("c:\libro1.xls", ListView1, ProgressBar1, 2)
'
'    If ret Then
'       ' -- OK
'       MsgBox " Datos exportados a Excel OK ", vbInformation
'    End If
'
'End Sub
'' ------------------------------------------------------------------
'' \\ -- Inicio
'' ------------------------------------------------------------------
'
'Private Sub Form_Load()
'
'    Dim lvItem  As ListItem
'    Dim i       As Integer
'
'    ' -- Configurar ListView
'    With ListView1
'        .ColumnHeaders.Add , , " Columna 1 "
'        .ColumnHeaders.Add , , " Columna 2 "
'        .View = lvwReport
'    End With
'
'    ' -- Añadir algunos items al listview
'    For i = 0 To 5000
'        Set lvItem = ListView1.ListItems.Add(, , " Item " & i)
'            lvItem.SubItems(1) = "subitem " & i
'    Next
'
'    Command1.Caption = " > Exportar Listview a Excel "
'
'End Sub
'' ------------------------------------------------------------------
'' \\ -- Redimensionar controles
'' ------------------------------------------------------------------
'Private Sub Form_Resize()
'    With ListView1
'        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - (Command1.Height + 50)
'        Command1.Move Me.ScaleWidth - (Command1.Width + 50), .Top + .Height + 50
'    End With
'    With Command1
'        ProgressBar1.Move 0, .Top, Me.ScaleWidth - (.Width + 100), .Height
'    End With
'End Sub
'
'
'
'
'Option Explicit
'
'' Importante : Agregar la referencia a Micorosft Excel xx object library
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Private Sub Form_Load()
'
'
'    'Variable de tipo Aplicación de Excel
'    Dim objExcel As Excel.Application
'
'    'Una variable de tipo Libro de Excel
'    Dim xLibro As Excel.Workbook
'    Dim Col As Integer, Fila As Integer
'
'    'creamos un nuevo objeto excel
'    Set objExcel = New Excel.Application
'
'    'Usamos el método open para abrir el archivo que está _
'     en el directorio del programa llamado archivo.xls
'    Set xLibro = objExcel.Workbooks.Open(App.Path + "\archivo.xls")
'
'    'Hacemos el Excel Visible
'    objExcel.visible = True
'
'    With xLibro
'
'        ' Hacemos referencia a la Hoja
'        With .Sheets(1)
'
'            'Recorremos la fila desde la 1 hasta la 7
'            For Fila = 1 To 7
'
'                'Agregamos el valor de la fila que _
'                 corresponde a la columna 2
'                Combo1.AddItem .Cells(Fila, 2)
'            Next
'
'        End With
'    End With
'
'    'Eliminamos los objetos si ya no los usamos
'    Set objExcel = Nothing
'    Set xLibro = Nothing
'
'End Sub
