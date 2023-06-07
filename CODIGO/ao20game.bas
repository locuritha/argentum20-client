Attribute VB_Name = "ao20game"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2023 - Noland Studios
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
Option Explicit
Option Base 0

Public Function InitGame() As Boolean
On Error GoTo InitGame_Err

    Set FormParser = New clsCursor
    Call FormParser.Init
    Call load_game_settings
    Call CheckResources
    If PantallaCompleta Then
        Call Resolution.SetResolution
    End If
    
   Exit Function

InitGame_Err:
    If Err.Number = 339 Then
        RegisterCom
    End If
    
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.Main", Erl)
    Resume Next
    
End Function

Public Sub MainLoop()
On Error GoTo GameLoop_Err
    DoEvents
    Do While prgRun
        Call ao20rendering.renderer.render
        prgRun = False
    Loop

    EngineRun = False
    
    Exit Sub
GameLoop_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Start", Erl)
    Resume Next
End Sub



