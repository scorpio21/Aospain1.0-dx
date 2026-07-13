Attribute VB_Name = "DibujarInventario"
Option Explicit

'[CODE]:MatuX
'
'
'[END]

Dim r1 As RECT
Dim r2 As RECT
Dim auxr As RECT
Dim mx As Integer, my As Integer
Dim bStaticInit As Boolean
Dim iFrameMod As Integer
Public OffsetDelInv As Integer

Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

Public Sub InitItem(lPreItem As Integer)
If lPreItem <= MAX_INVENTORY_SLOTS Then _
If UserInventory(lPreItem).GrhIndex > 0 Then _
    ItemElegido = lPreItem: bInvMod = True
End Sub

Private Sub InitMem()
    r1.Right = 32: r1.Bottom = 32
    r2.Right = 32: r2.Bottom = 32
    auxr.Right = 32: auxr.Bottom = 32
    iFrameMod = 1
    bStaticInit = True
End Sub

'[CODE]:MatuX
Public Sub DibujarInvBox()
    ' Item seleccionado en el slot - sin implementacion DX8 especial
End Sub
'[END]

Sub DibujarInv()
' [DX8] El inventario usa GDI directo al picInv (NO engine.Device_Box que va al renderer)

Dim iX As Integer
Dim destX As Integer, destY As Integer
Dim col As Integer, row As Integer

If Not bStaticInit Then Call InitMem

frmMain.picInv.Cls

col = 0
row = 0

For iX = OffsetDelInv + 1 To UBound(UserInventory)
    If UserInventory(iX).GrhIndex > 0 Then
        destX = col * 32
        destY = row * 32
        
        ' Dibujar grafico via GDI
        Call GrhRenderToHdc(UserInventory(iX).GrhIndex, frmMain.picInv.hdc, destX, destY, True)
        
        ' Cantidad
        If UserInventory(iX).Amount > 1 Then
            frmMain.picInv.ForeColor = vbWhite
            frmMain.picInv.CurrentX = destX
            frmMain.picInv.CurrentY = destY + 18
            frmMain.picInv.Print CStr(UserInventory(iX).Amount)
        End If
        
        ' Equipado
        If UserInventory(iX).Equipped Then
            frmMain.picInv.ForeColor = vbYellow
            frmMain.picInv.CurrentX = destX + 18
            frmMain.picInv.CurrentY = destY
            frmMain.picInv.Print "+"
        End If
        
        ' Seleccion
        If ItemElegido = iX Then
            frmMain.picInv.Line (destX, destY)-(destX + 31, destY + 31), vbYellow, B
        End If
    End If
    
    col = col + 1
    If col >= 4 Then
        col = 0
        row = row + 1
    End If
    
    If row * 32 > frmMain.picInv.ScaleHeight Then Exit For
Next iX

frmMain.picInv.Refresh
End Sub
