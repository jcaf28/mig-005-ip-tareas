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

    ' === 5. Importar datos desde Excel utilizando TransferSpreadsheet ===
    logStr = ""
    
    ' NUEVAS IMPORTACIONES (ORDEN ESPECÍFICO)
    ' Importar procesos
    logStr = logStr & ">> Insertando T_PROCESOS_SUBIR.xlsx en T_PROCESOS..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_PROCESOS_SUBIR.xlsx", "T_PROCESOS", logStr)
    
    ' Importar grupos
    logStr = logStr & vbCrLf & ">> Insertando TJ_GRUPOS_SUBIR.xlsx en TJ_GRUPOS..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\TJ_GRUPOS_SUBIR.xlsx", "TJ_GRUPOS", logStr)
    
    ' Importar tareas
    logStr = logStr & vbCrLf & ">> Insertando T_TAREAS_SUBIR.xlsx en T_TAREAS..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_TAREAS_SUBIR.xlsx", "T_TAREAS", logStr)
    
    ' IMPORTACIONES ORIGINALES
    ' Importar obras
    logStr = logStr & vbCrLf & ">> Insertando T_OBRAS_SUBIR.xlsx en T_OBRAS..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_OBRAS_SUBIR.xlsx", "T_OBRAS", logStr)

    ' Importar usuarios
    logStr = logStr & vbCrLf & ">> Insertando T_USUARIOS_SUBIR.xlsx en T_USUARIOS..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_USUARIOS_SUBIR.xlsx", "T_USUARIOS", logStr)
    
    ' Importar anotaciones
    logStr = logStr & vbCrLf & ">> Insertando T_ANOTACIONES_SUBIR.xlsx en T_ANOTACIONES..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_ANOTACIONES_SUBIR.xlsx", "T_ANOTACIONES", logStr)
    
    ' Importar anotaciones validadas
    logStr = logStr & vbCrLf & ">> Insertando T_ANOTACIONES_VALID_SUBIR.xlsx en T_ANOTACIONES_VALID..." & vbCrLf
    Call ImportarDesdeExcelRapido(ARCHIVOS_PATH & "\T_ANOTACIONES_VALID_SUBIR.xlsx", "T_ANOTACIONES_VALID", logStr)

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

Private Sub ImportarDesdeExcelRapido(rutaExcel As String, nombreTablaDestino As String, ByRef logStr As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tablaTemporal As String
    Dim contador As Long
    Dim totalRegistros As Long
    Dim registrosInsertados As Long
    Dim sqlCount As String
    Dim sqlAppend As String
    
    Set db = CurrentDb
    
    ' Crear nombre para tabla temporal
    tablaTemporal = "Temp_" & Format(Now, "yyyymmddhhnnss")
    
    ' Importar Excel completo a tabla temporal
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tablaTemporal, rutaExcel, True
    
    ' Contar registros en tabla temporal
    sqlCount = "SELECT COUNT(*) FROM [" & tablaTemporal & "];"
    totalRegistros = DCount("*", tablaTemporal)
    
    ' Preparar consulta de anexado para transferir datos de tabla temporal a tabla destino
    sqlAppend = "INSERT INTO [" & nombreTablaDestino & "] " & _
                "SELECT * FROM [" & tablaTemporal & "];"
                
    ' Intentar insertar todos los registros de una vez
    On Error Resume Next
    db.Execute sqlAppend, dbFailOnError
    
    If Err.Number = 0 Then
        ' Si no hay error, todos se insertaron correctamente
        registrosInsertados = totalRegistros
        logStr = logStr & "   Todos los registros insertados correctamente." & vbCrLf
        logStr = logStr & "   Total registros: " & totalRegistros & vbCrLf
    Else
        ' Si hubo error, intentar insertar registro por registro para registrar los que fallan
        logStr = logStr & "   Error al insertar en bloque: " & Err.Description & vbCrLf
        logStr = logStr & "   Intentando inserción registro por registro..." & vbCrLf
        
        Dim rstOrigen As DAO.Recordset
        Dim rstDestino As DAO.Recordset
        Dim i As Integer
        
        registrosInsertados = 0
        Set rstOrigen = db.OpenRecordset(tablaTemporal)
        Set rstDestino = db.OpenRecordset(nombreTablaDestino)
        
        ' Obtener array de nombres de campo
        Dim fields() As String
        ReDim fields(rstOrigen.Fields.Count - 1)
        For i = 0 To rstOrigen.Fields.Count - 1
            fields(i) = rstOrigen.Fields(i).Name
        Next i
        
        ' Intentar insertar cada registro individualmente
        contador = 0
        Do While Not rstOrigen.EOF
            contador = contador + 1
            On Error Resume Next
            rstDestino.AddNew
            
            For i = 0 To rstOrigen.Fields.Count - 1
                rstDestino(fields(i)).Value = rstOrigen(fields(i)).Value
            Next i
            
            rstDestino.Update
            
            If Err.Number = 0 Then
                registrosInsertados = registrosInsertados + 1
            Else
                logStr = logStr & "      Error en registro " & contador & ": " & Err.Description & vbCrLf
                Err.Clear
            End If
            
            rstOrigen.MoveNext
        Loop
        
        rstOrigen.Close
        rstDestino.Close
    End If
    
    On Error GoTo ErrorHandler
    
    ' Eliminar tabla temporal
    DoCmd.DeleteObject acTable, tablaTemporal
    
    logStr = logStr & "   Total registros: " & totalRegistros & vbCrLf
    logStr = logStr & "   Registros insertados: " & registrosInsertados & vbCrLf
    logStr = logStr & "   Registros fallidos: " & (totalRegistros - registrosInsertados) & vbCrLf
    
    Exit Sub
    
ErrorHandler:
    logStr = logStr & "   Error al importar: " & Err.Number & " - " & Err.Description & vbCrLf
    
    ' Intentar eliminar la tabla temporal si existe
    On Error Resume Next
    DoCmd.DeleteObject acTable, tablaTemporal
End Sub

Private Sub EscribirLog(texto As String, ruta As String)
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(ruta, True, True) ' True = overwrite, True = UTF-8
    f.Write texto
    f.Close
End Sub