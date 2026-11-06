Attribute VB_Name = "Module1"
Option Explicit
'-Indispensable para browse del combo
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Long) As Long

'-Sonido Beep   frecuencia, Duracion
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


Public cnx8  As ADODB.Connection
Public cDSource8, cICatal8, CString8, cUsr8, cPwd8 As String
Public cnx  As ADODB.Connection
Public cDSource, cICatal, CString, cUsr, cPwd As String
' Variable global para la conexión (ajusta según tu proyecto)
' Private gCn As ADODB.Connection
Public gIdSesion As String
Public wtpoquery, wacceso1, wacceso2, wacceso3, wacceso4, wacceso5, wacceso6, encclave, cNamedbf, wNombFile, carpeta As String
Public wacceso As Boolean
Public wrazonsocial, wsistema As String


Public Sub WVencripta(cclave As String)
    Dim i, wvlen As Integer
    Dim wchr, wvasc As String
    wvlen = Len(Trim(cclave))
    encclave = ""
    
    'f_show = False
    'For i = Forms.Count - 1 To 0 Step -1
    '  If Forms(i).Visible = False Then
    '  Else
        '- Form Abierto
        'If Forms(i).Name = "CorrigeEGR" Then
        '   f_show = True
        'End If
      'End If
    'Next
    
    For i = 1 To wvlen
       wchr = Mid(cclave, i, 1)
       wvasc = Asc(wchr) + (10 + i)
         'If f_show Then
         '   MsgBox Trim(Str(i)) + ":" + wchr + "->" + Trim(Str(wvasc)) + "[" + Chr(wvasc) + "]"
         'End If
       encclave = encclave + Chr(wvasc)
    Next
End Sub
Public Sub WVdesencripta(cclave As String)
    Dim i, wvlen As Integer
    Dim wchr, wvasc As String
    wvlen = Len(Trim(cclave))
    encclave = ""
    
    'f_show = False
    'For i = Forms.Count - 1 To 0 Step -1
    '  If Forms(i).Visible = False Then
    '  Else
        '- Form Abierto
        'If Forms(i).Name = "CorrigeEGR" Then
        '   f_show = True
        'End If
      'End If
    'Next
    
    For i = 1 To wvlen
       wchr = Mid(cclave, i, 1)
       wvasc = Asc(wchr) - (10 + i)
     '    If f_show Then
     '       MsgBox Trim(Str(i)) + ":" + wchr + "->" + Trim(Str(wvasc)) + "[" + Chr(wvasc) + "]"
     '    End If
       encclave = encclave + Chr(wvasc)
    Next
End Sub


'-Export using Microsoft Excel Object Library, This code copies data from temp table to an excel file(temp.xls),
'-delete the generated xml file at the end of the program as its just an temp file.
'-Add a project reference to ADODB & Excel Object library before using this code,
'-crear tabla temporal (nombre varchar (10), int edad) create table temp ( name varchar(10), age int)
'-insertar en los valores temporales ('XYZ', 20) insert into temp values ( 'xyz', 20)
'-insertar en los valores temporales ('ABC', 30) insert into temp values ( 'abc', 30)
Public Sub ExportTableToXLS_click()        '    wTpoQuery As String  **** OUT ***
'==================================
 On Error GoTo Control
 Dim SqlXls As String
 Dim strExcelFile, strXmlFile As String
 'Dim cnx As ADODB.Connection
 Dim Rs As ADODB.Recordset
 Dim eApp As Excel.Application
 Dim wk As New Excel.Workbook
 'Set cnx = New ADODB.Connection
 Set Rs = New ADODB.Recordset
 
 CreaDir ("c:\rpt")
 
 '-Si ficheros excel y xm; existen, aqui se borran
 Select Case wtpoquery
    Case "0"
       strExcelFile = cNamedbf      ' "C:\RPT\QUERYSOP.XLS"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\OBQMULTIPLE10.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #Tabla"
       wNombFile = "OBQMULTIPLE10"
    Case "1"
       strExcelFile = cNamedbf      ' "C:\RPT\QUERYPROG.XLS"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\OBQMULTIPLE.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #Tabla1"
       wNombFile = "OBQMULTIPLE"
    Case "2"
    Case "3"
 End Select
 
 'cnx.Open "Provider=SQLOLEDB.1;Password=password;Persist Security Info=True;User ID=sa;Initial Catalog=test;Data Source=TDL06B21"
 'rs.Open "SELECT * FROM tablatmp", Cnx, adOpenStatic
 Rs.Open SqlXls, cnx, adOpenStatic
 Rs.Save strXmlFile, adPersistXML
 Set eApp = New Excel.Application
 Set wk = eApp.Workbooks.Add
 With wk.ActiveSheet.QueryTables.Add(Connection:="FINDER;" & strXmlFile, Destination:=wk.ActiveSheet.Range("A1"))
   .Name = wNombFile
   .FieldNames = True
   .RowNumbers = False
   .FillAdjacentFormulas = False
   .PreserveFormatting = True
   .RefreshOnFileOpen = False
   .BackgroundQuery = True
   .RefreshStyle = xlInsertDeleteCells
   .SavePassword = False
   .SaveData = True
   .AdjustColumnWidth = True
   .RefreshPeriod = 0
   .WebSelectionType = xlAllTables
   .WebFormatting = xlWebFormattingNone
   .WebPreFormattedTextToColumns = True
   .WebConsecutiveDelimitersAsOne = True
   .WebSingleBlockTextImport = False
   .WebDisableDateRecognition = False
   .WebDisableRedirections = False
   .Refresh BackgroundQuery:=False
  End With
  wk.SaveAs strExcelFile
  wk.Close
  eApp.Quit
   Exit Sub
Control:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr
End Sub


Public Function Exportar_ADO_Excel()
      On Error GoTo Control
        '-Dim cn          As New ADODB.Connection
        Dim rec         As New ADODB.Recordset
        Dim Excel       As Object
        Dim Libro       As Object
        Dim Hoja        As Object
        Dim ArrData     As Variant
        Dim iRec        As Long
        Dim iCol        As Integer
        Dim iRow        As Integer
                
        Dim sqlTabla, strExcelFile      As String
        
        Select Case wtpoquery
        Case "0"
          strExcelFile = cNamedbf
          sqlTabla = "SELECT * FROM #Tabla"
          wNombFile = "OBQMULTIPLE10"
        Case "1"
          strExcelFile = cNamedbf
          sqlTabla = "SELECT * FROM #Tabla1"
          wNombFile = "OBQMULTIPLE"
        Case "1.5"
          strExcelFile = cNamedbf
          sqlTabla = "SELECT * FROM #TablaC"
          wNombFile = "OBProducDia"
        Case "3"
        End Select
        
        
        '-Si fichero excel existe, aqui es borrado
        If Dir(strExcelFile) <> "" Then Kill strExcelFile
       
        'Exportar_ADO_Excel(sPathDB, Consulta, "c:\libro.xLS")
        Set rec = New ADODB.Recordset
        ' -- Abrir el Recordset pasándole la cadena sql
        rec.Open sqlTabla, cnx, adOpenStatic
        ' -- Crear los objetos para utilizar el Excel
        Set Excel = CreateObject("Excel.Application")
        Set Libro = Excel.Workbooks.Add
          
        ' -- Hacer referencia a la hoja
        Set Hoja = Libro.Worksheets(1)
          
        Excel.Visible = True: Excel.UserControl = True
        iCol = rec.Fields.Count
        For iCol = 1 To rec.Fields.Count
            Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
        Next
          
        If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
            Hoja.Cells(2, 1).CopyFromRecordset rec
        Else
      
            ArrData = rec.GetRows
      
            iRec = UBound(ArrData, 2) + 1
              
            For iCol = 0 To rec.Fields.Count - 1
                For iRow = 0 To iRec - 1
      
                    If IsDate(ArrData(iCol, iRow)) Then
                        ArrData(iCol, iRow) = Format(ArrData(iCol, iRow))
      
                    ElseIf IsArray(ArrData(iCol, iRow)) Then
                        ArrData(iCol, iRow) = "Array Field"
                    End If
                Next iRow
            Next iCol
                  
            ' -- pasa los datos a la hoja de Excel
            Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(ArrData)
        End If
      
        Excel.Selection.CurrentRegion.Columns.AutoFit
        Excel.Selection.CurrentRegion.Rows.AutoFit
      
        ' -- Cierra el recordset y la base de datos y los objetos ADO
        rec.Close
        'Cnx.Close
          
        Set rec = Nothing
        'Set Cnx = Nothing
        ' -- graba el libro XLS
        Libro.SaveAs strExcelFile
        Libro.Close
        ' -- Elimina las referencias Xls
        Set Hoja = Nothing
        Set Libro = Nothing
        Excel.Quit
        Set Excel = Nothing
    
      Exit Function
Control:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr

End Function

 Public Function GetData(vValue As Variant) As Variant
        Dim x As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
          
        xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
          
        ReDim T(xMax, yMax)
        For x = 0 To xMax
            For Y = 0 To yMax
                T(x, Y) = vValue(Y, x)
            Next Y
        Next x
          
        GetData = T
 End Function

Public Function CreaDir(valdir As String)
   carpeta = Dir(valdir & "\", vbDirectory)    'comprueba si ya existe la carpeta "fautos"
   If carpeta = "" Then
      MkDir valdir
      MsgBox "Carpeta creada " & valdir
   End If
  
  ' Dim AttrDev%
  ' On Error Resume Next
  ' AttrDev = GetAttr(valdir)
  ' If Err.Number Then
  '    Err.Clear
  '    MkDir valdir
  '    MsgBox "Creo Dir " & valdir
  '  End If
End Function



