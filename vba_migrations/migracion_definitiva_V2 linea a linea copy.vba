Option Compare Database
Option Explicit

Private Const BASE_PATH As String = "C:\Users\18287\projects\powerBi\PowerBI ip-imputaciones\prod\migraciones\mig-005-IP-tareas"
Private Const ARCHIVOS_PATH As String = BASE_PATH & "\archivos\archivos definitivos migracion"
Private Const LOG_PATH As String = BASE_PATH & "\vba_migrations\logs"

Private Sub EstablecerCamposRequeridos()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim camposRequeridos As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    Set tdf = db.TableDefs("T_ANOTACIONES")
    
    ' Lista de campos que queremos establecer como requeridos
    camposRequeridos = Array("ClaveObra", "CodTarea", "Idusuario")
    
    ' Establecer cada campo como requerido
    For i = LBound(camposRequeridos) To UBound(camposRequeridos)
        Set fld = tdf.Fields(camposRequeridos(i))
        fld.Required = True
    Next i
    
    db.TableDefs.Refresh
End Sub

Private Sub EliminarTablasOld(db As DAO.Database)
    Dim rst As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT Name " & _
          "FROM MSysObjects " & _
          "WHERE Type=1 " & _
          "  AND Name Like '*_OLD' " & _
          "  AND Name NOT LIKE 'MSys*';"
    
    Set rst = db.OpenRecordset(sql)
    Do While Not rst.EOF
        db.Execute "DROP TABLE [" & rst!Name & "]", dbFailOnError
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub MigracionUsuariosAnotaciones()
    Const REL_NAME As String = "rel_USUARIOS_ANOTACIONES"
    Const REL_VALID_OBRAS As String = "rel_VALID_OBRAS"
    
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim listaUsuarios As String
    Dim totalIntentados As Long, totalInsertados As Long
    Dim logStr As String, timestamp As String, logFile As String

    Set db = CurrentDb
    On Error GoTo ErrControl

    ' === 0. Limpieza previa y establecer campos requeridos ===
    Call EliminarTablasOld(db)
    Call EstablecerCamposRequeridos

    ' === 1. Copias de seguridad ===
    db.Execute "SELECT * INTO T_ANOTACIONES_OLD FROM T_ANOTACIONES;", dbFailOnError
    db.Execute "SELECT * INTO T_ANOTACIONES_VALID_OLD FROM T_ANOTACIONES_VALID;", dbFailOnError

    ' === 2. Normalización ===
    db.Execute "UPDATE T_ANOTACIONES SET Idusuario='0003' WHERE Idusuario='3';", dbFailOnError

    ' === 3. Eliminación de usuarios no válidos ===
    listaUsuarios = "'carraiza','14859','14720','18287','14550','14902'"
    
    db.Execute _
        "DELETE T_ANOTACIONES_VALID.* " & _
        "FROM T_ANOTACIONES_VALID " & _
        "INNER JOIN T_ANOTACIONES " & _
        "  ON T_ANOTACIONES_VALID.IdAnot = T_ANOTACIONES.IdAnot " & _
        "WHERE T_ANOTACIONES.Idusuario IN (" & listaUsuarios & ");", _
        dbFailOnError

    db.Execute "DELETE FROM T_ANOTACIONES WHERE Idusuario IN (" & listaUsuarios & ");", dbFailOnError

    ' === 4. Crear relación RI entre USUARIOS y ANOTACIONES ===
    On Error Resume Next
    db.Relations.Delete REL_NAME
    On Error GoTo ErrControl
    
    Set rel = db.CreateRelation(REL_NAME, "T_USUARIOS", "T_ANOTACIONES", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("IdUsuario")
    fld.ForeignName = "Idusuario"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' === 4b. Crear relación RI entre OBRAS y ANOTACIONES_VALID ===
    On Error Resume Next
    db.Relations.Delete REL_VALID_OBRAS
    On Error GoTo ErrControl
    
    Set rel = db.CreateRelation(REL_VALID_OBRAS, "T_OBRAS", "T_ANOTACIONES_VALID", _
                               dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("ClaveObra")
    fld.ForeignName = "ClaveObra"
    rel.Fields.Append fld
    db.Relations.Append rel

    ' === 5. Importar datos desde Excel ===
    logStr = ""
    totalIntentados = 0: totalInsertados = 0
    
    ' Importar obras
    logStr = logStr & ">> Insertando T_OBRAS_SUBIR.xlsx en T_OBRAS..." & vbCrLf
    Call ImportarDesdeExcel(ARCHIVOS_PATH & "\T_OBRAS_SUBIR.xlsx", "T_OBRAS", logStr, totalIntentados, totalInsertados)

    ' Importar usuarios
    logStr = logStr & vbCrLf & ">> Insertando T_USUARIOS_SUBIR.xlsx en T_USUARIOS..." & vbCrLf
    Call ImportarDesdeExcel(ARCHIVOS_PATH & "\T_USUARIOS_SUBIR.xlsx", "T_USUARIOS", logStr, totalIntentados, totalInsertados)
    
    ' Importar anotaciones
    logStr = logStr & vbCrLf & ">> Insertando T_ANOTACIONES_SUBIR.xlsx en T_ANOTACIONES..." & vbCrLf
    Call ImportarDesdeExcel(ARCHIVOS_PATH & "\T_ANOTACIONES_SUBIR.xlsx", "T_ANOTACIONES", logStr, totalIntentados, totalInsertados)
    
    ' Importar anotaciones validadas
    logStr = logStr & vbCrLf & ">> Insertando T_ANOTACIONES_VALID_SUBIR.xlsx en T_ANOTACIONES_VALID..." & vbCrLf
    Call ImportarDesdeExcel(ARCHIVOS_PATH & "\T_ANOTACIONES_VALID_SUBIR.xlsx", "T_ANOTACIONES_VALID", logStr, totalIntentados, totalInsertados)

    ' === 6. Guardar log ===
    timestamp = Format(Now, "yyyymmdd_HHMMSS")
    logFile = LOG_PATH & "\migration_" & timestamp & ".txt"
    Call EscribirLog(logStr, logFile)

    ' === 7. Limpieza final de *_OLD (opcional) ===
    Call EliminarTablasOld(db)

    MsgBox "Migración completada correctamente." & vbCrLf & _
           "Log generado: " & logFile, vbInformation

    Salida:
        Set fld = Nothing
        Set rel = Nothing
        Set db = Nothing
        Exit Sub

    ErrControl:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        Resume Salida
End Sub

Private Sub ImportarDesdeExcel(rutaExcel As String, nombreTablaDestino As String, _
                               ByRef logStr As String, ByRef totalIntentados As Long, ByRef totalInsertados As Long)
    On Error GoTo ErrHandlerGeneral

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim rst As DAO.Recordset
    Dim fila As Long, col As Long
    Dim headers() As String
    Dim campos As Long
    Dim insertados As Long
    Dim fallidos As Long

    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Open(rutaExcel, False, True)
    Set xlWs = xlWb.Sheets(1)

    campos = xlWs.UsedRange.Columns.Count
    ReDim headers(1 To campos)
    For col = 1 To campos
        headers(col) = xlWs.Cells(1, col).Value
    Next col

    Set rst = CurrentDb.OpenRecordset(nombreTablaDestino, dbOpenDynaset)
    fila = 2
    insertados = 0
    fallidos = 0

    Do While xlWs.Cells(fila, 1).Value <> ""
        On Error Resume Next
        rst.AddNew
        For col = 1 To campos
            rst(headers(col)).Value = xlWs.Cells(fila, col).Value
        Next col
        rst.Update
        If Err.Number = 0 Then
            insertados = insertados + 1
        Else
            fallidos = fallidos + 1
            logStr = logStr & "      Línea " & fila & ": " & Err.Description & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
        fila = fila + 1
    Loop

    rst.Close
    xlWb.Close False
    xlApp.Quit

    logStr = logStr & "   Registros intentados: " & (fila - 2) & vbCrLf
    logStr = logStr & "   Registros insertados: " & insertados & vbCrLf
    logStr = logStr & "   Registros fallidos: " & fallidos & vbCrLf

    totalIntentados = totalIntentados + (fila - 2)
    totalInsertados = totalInsertados + insertados

    Set rst = Nothing
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Exit Sub

    ErrHandlerGeneral:
        logStr = logStr & "   ? Error general al importar: " & Err.Description & vbCrLf
        On Error Resume Next
        If Not xlApp Is Nothing Then xlApp.Quit
        Set xlApp = Nothing
End Sub


Private Sub EscribirLog(texto As String, ruta As String)
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(ruta, True, True) ' True = overwrite, True = UTF-8
    f.Write texto
    f.Close
End Sub




