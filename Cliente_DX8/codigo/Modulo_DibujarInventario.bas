Attribute VB_Name = "DibujarInventario"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mataas Fernando Pequeao
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 namero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cadigo Postal 1900
'Pablo Ignacio Marquez



Option Explicit

'[CODE]:MatuX
'
'  Casi todo recodeado menos los calculos
'
'[END]'

Public Const XCantItems = 5

Public OffsetDelInv As Long
Public ItemElegido As Long
Public mx As Long
Public my As Long

' [DX8] Superficies eliminadas - renderizado via engine
Private bStaticInit  As Boolean   'Se inicializaron las Statics?
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT  'Pos del cuadradito rojo
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer


Function ClicEnItemElegido(x As Integer, Y As Integer) As Boolean
bInvMod = True
mx = x \ 32 + 1
my = Y \ 32 + 1
If ItemElegido = 0 Or FLAGORO Then
    ClicEnItemElegido = False
Else
    ClicEnItemElegido = (UserInventory(ItemElegido).OBJIndex > 0) And (ItemElegido = (mx + (my - 1) * 5) + OffsetDelInv)
End If
End Function

Sub ItemClick(x As Integer, Y As Integer)
Dim lPreItem As Long

bInvMod = False
mx = x \ 32 + 1
my = Y \ 32 + 1

lPreItem = (mx + (my - 1) * 5) + OffsetDelInv

If lPreItem <= MAX_INVENTORY_SLOTS Then _
If UserInventory(lPreItem).GrhIndex > 0 Then _
    ItemElegido = lPreItem: bInvMod = True
End Sub

'[CODE]:MatuX'
Public Sub DibujarInvBox()
    On Error Resume Next
    If bStaticInit And ItemElegido <> 0 Then
        ' [DX8] BltColorFill eliminado - engine limpia el frame con Clear en BeginFrame
        ' [DX8] SelSurface eliminada - no se necesita copia de surface a surface en DX8
        ' El renderizado directo via engine reemplaza este patron
        
        'inline hardcoded DDrawTransGrhToSurface()
       '{
        With Grh(1)
            '.FrameCounter = .FrameCounter + iFrameMod
            'If (.FrameCounter >= GrhData(.GrhIndex).NumFrames) Then
            '    iFrameMod = -1
            'ElseIf (.FrameCounter <= 1) Then
            '    iFrameMod = 1
            'End If
            .FrameCounter = 2
            ' [DX8] Box frame via engine
            Dim dxLightBox(0 To 3) As Long
            dxLightBox(0) = &HFFFFFFFF : dxLightBox(1) = &HFFFFFFFF
            dxLightBox(2) = &HFFFFFFFF : dxLightBox(3) = &HFFFFFFFF
            Dim boxGrh As Long : boxGrh = GrhData(.GrhIndex).Frames(.FrameCounter)
            Call engine.Device_Box_Textured_Render(boxGrh, 0, 0, GrhData(boxGrh).pixelWidth, GrhData(boxGrh).pixelHeight, dxLightBox(), GrhData(boxGrh).sX, GrhData(boxGrh).sY, False)
        End With
       '}
        ' [DX8] BltToDC eliminado - el render ya se hizo arriba via engine.Device_Box_Textured_Render
        Call frmMain.picInv.Refresh
    End If
End Sub
'[END]'

'Sub DibujarInv(PictureHandler As Long, desp As Integer)
Sub DibujarInv()
'[CODE]:MatuX'
'
'
'[END]'

Dim iX As Integer

If Not bStaticInit Then _
    Call InitMem

r1.Top = 0: r1.Left = 0: r1.Right = 32: r1.Bottom = 32
r2.Top = 0: r2.Left = 0: r2.Right = 32: r2.Bottom = 32

frmMain.picInv.Cls

For iX = OffsetDelInv + 1 To UBound(UserInventory)
    If UserInventory(iX).GrhIndex > 0 Then
        ' [DX8] BltColorFill eliminado
        ' [DX8] Aux item via engine
        Dim dxLightAux(0 To 3) As Long
        dxLightAux(0) = &HFFFFFFFF : dxLightAux(1) = &HFFFFFFFF
        dxLightAux(2) = &HFFFFFFFF : dxLightAux(3) = &HFFFFFFFF
        Dim auxGrh As Long : auxGrh = GrhData(UserInventory(iX).GrhIndex).FileNum
        Call engine.Device_Box_Textured_Render(auxGrh, 0, 0, GrhData(auxGrh).pixelWidth, GrhData(auxGrh).pixelHeight, dxLightAux(), GrhData(auxGrh).sX, GrhData(auxGrh).sY, False)
        ' [DX8] Texto cantidad via engine.Text_Render_ext
        Call engine.Text_Render_ext(CStr(UserInventory(iX).Amount), 0, 0, 32, 12, &HFFFFFFFF)

        If UserInventory(iX).Equipped Then
            Call engine.Text_Render_ext("+", 20, 20, 12, 12, &H00FFFF00)
        End If

        If ItemElegido = iX Then
            'Call SelSurface.BltColorFill(auxr, vbBlack)
            'Call SelSurface.BltFast(0, 0, AuxSurface, auxr, DDBLTFAST_SRCCOLORKEY)

            With r2: .Left = (mx - 1) * 32: .Right = r2.Left + 32: .Top = (my - 1) * 32: .Bottom = r2.Top + 32: End With
            'With rBox: .Top = r2.Top: .Left = r2.Left: .Bottom = r2.Bottom: .Right = r2.Right: End With
            
            ' [DX8] Sel frame via engine
            Dim dxLightSel(0 To 3) As Long
            dxLightSel(0) = &HFFFFFFFF : dxLightSel(1) = &HFFFFFFFF
            dxLightSel(2) = &HFFFFFFFF : dxLightSel(3) = &HFFFFFFFF
            Dim selGrh As Long : selGrh = GrhData(GrhData(Grh(1).GrhIndex).Frames(2)).FileNum
            Call engine.Device_Box_Textured_Render(selGrh, 0, 0, GrhData(selGrh).pixelWidth, GrhData(selGrh).pixelHeight, dxLightSel(), GrhData(selGrh).sX, GrhData(selGrh).sY, False)
        End If
        ' [DX8] BltToDC eliminado - el render ya se hizo arriba via engine.Device_Box_Textured_Render
    End If

    r2.Left = r2.Left + 32
    r2.Right = r2.Right + 32
    r1.Left = r1.Left + 32
    r1.Right = r1.Right + 32
    If r2.Left >= 160 Then
        r2.Left = 0
        r1.Left = 0
        r1.Right = 32
        r2.Right = 32
        r2.Top = r2.Top + 32
        r1.Top = r1.Top + 32
        r2.Bottom = r2.Bottom + 32
        r1.Bottom = r1.Bottom + 32
    End If
Next iX

'frmMain.picInv.Refresh

'Call DibujarInvBox

bInvMod = False

If ItemElegido = 0 Then _
    Call ItemClick(2, 2)

End Sub

Private Sub InitMem()
    'Back Buffer Surface
    r1.Right = 32: r1.Bottom = 32
    r2.Right = 32: r2.Bottom = 32

    ' [DX8] Superficies, SurfaceDesc, ColorKey y fuentes de DirectDraw ya no se usan.
    ' El texto y las superficies se manejan via engine (Direct3D8 + D3DX8 font)

    auxr.Right = 32: auxr.Bottom = 32

    'RedBox Frame Position List
    With rBoxFrame(0): .Left = 0:  .Top = 0: .Right = 32: .Bottom = 32: End With
    With rBoxFrame(1): .Left = 32: .Top = 0: .Right = 64: .Bottom = 32: End With
    With rBoxFrame(2): .Left = 64: .Top = 0: .Right = 96: .Bottom = 32: End With
    iFrameMod = 1

    bStaticInit = True
End Sub

