VERSION 5.00
Begin VB.Form Frmcarga 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cargando..."
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Frmcarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadInterface("VentanaCargando.bmp")
    MakeFormTransparent Me, vbBlack

    VerificarMD5 'Verificamos archivos vitales ENFERMiTO

    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "Frmcarga.Form_Load", Erl)
    Resume Next
    
End Sub
'Verificamos archivos vitales ENFERMiTO
Private Function MD5Hash(ByVal sInput As String) As String
    Dim md5 As Object
    Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Dim Bytes() As Byte
    Bytes = StrConv(sInput, vbFromUnicode)
    
    Dim Hash() As Byte
    Hash = md5.ComputeHash_2((Bytes))
    
    Dim i As Integer
    Dim sOutput As String
    sOutput = ""
    For i = LBound(Hash) To UBound(Hash)
        sOutput = sOutput & Right("0" & hex(Hash(i)), 2)
    Next i
    
    MD5Hash = sOutput
End Function

Private Function LeerInt(ByVal Ruta As String) As Integer
f = FreeFile
Open Ruta For Input As f
LeerInt = Input$(LOF(f), #f)
Close #f
End Function


Private Function LoadFile(ByVal sFilePath As String) As String
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open sFilePath For Binary Access Read As FileNumber
    LoadFile = Space$(LOF(FileNumber))
    Get FileNumber, , LoadFile
    Close FileNumber
End Function

Sub VerificarMD5()
    Dim appPath As String
    appPath = App.path
    
    Dim graficosMD5 As String
    graficosMD5 = "4B331596090809AA9AA883DADEF89CB4" ' graficos
    
    Dim initMD5 As String
    initMD5 = "3425BC278AD7D7B29865E63935EE9D7F" ' init
    
    Dim mapasMD5 As String
    mapasMD5 = "BB60ED37C88511B2B579D777FD1DA5F9" ' mapas
    
   
    Dim graficosFile As String
    graficosFile = appPath & "\..\Recursos\OUTPUT\Graficos"
    
    Dim initFile As String
    initFile = appPath & "\..\Recursos\OUTPUT\Init"
    
    Dim mapasFile As String
    mapasFile = appPath & "\..\Recursos\OUTPUT\Mapas"
    

    
  
    If MD5Hash(LoadFile(graficosFile)) = graficosMD5 Then
        'MsgBox "Ok: El archivo graficos tiene el MD5 correcto.", vbInformation
    Else
        MsgBox "El archivo de graficos ha sido alterado. Reinstale el cliente.", vbExclamation
        End
    End If
    
    If MD5Hash(LoadFile(initFile)) = initMD5 Then
        'MsgBox "Ok: El archivo init tiene el MD5 correcto.", vbInformation
    Else
        MsgBox "El archivo de inits ha sido alterado. Reinstale el cliente.", vbExclamation
        End
    End If
    
        If MD5Hash(LoadFile(mapasFile)) = mapasMD5 Then
        'MsgBox "Ok: El archivo mapas tiene el MD5 correcto.", vbInformation
    Else
        MsgBox "El archivo de mapas ha sido alterado. Reinstale el cliente.", vbExclamation
        End
    End If
    

End Sub
'Verificamos archivos vitales ENFERMiTO
