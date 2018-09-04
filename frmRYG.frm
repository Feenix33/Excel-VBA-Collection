VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRYG 
   Caption         =   "Select Rule Values"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3465
   OleObjectBlob   =   "frmRYG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRYG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iValG As Long
Dim iValY As Long
Dim iValR As Long
Dim iClrG As Long
Dim iClrY As Long
Dim iClrR As Long

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Function fcnColor(clrIn) As Long
    Dim clrR As Integer
    Dim clrG As Integer
    Dim clrB As Integer
    
    clrR = clrIn Mod 256
    clrG = (clrIn \ 256) Mod 256
    clrB = clrIn \ 65536
    
    If Application.Dialogs(xlDialogEditColor).Show(1, clrR, clrG, clrB) = True Then
        clrIn = ActiveWorkbook.Colors(1)
    Else
        '
    End If
    fcnColor = clrIn
End Function
Private Sub btnG_Click()
    iClrG = fcnColor(iClrG)
    btnG.BackColor = iClrG
End Sub
Private Sub btnY_Click()
    iClrY = fcnColor(iClrY)
    btnY.BackColor = iClrY
End Sub
Private Sub btnR_Click()
    iClrR = fcnColor(iClrR)
    btnR.BackColor = iClrR
End Sub

Private Sub btnOK_Click()
    iHeatMapG = CInt(txtG.Text)
    iHeatMapY = CInt(txtY.Text)
    iHeatMapR = CInt(txtR.Text)
    iHeatClrG = btnG.BackColor
    iHeatClrY = btnY.BackColor
    iHeatClrR = btnR.BackColor
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    iValG = iHeatMapG
    iValY = iHeatMapY
    iValR = iHeatMapR
    iClrG = iHeatClrG
    iClrY = iHeatClrY
    iClrR = iHeatClrR
    
    btnG.BackColor = iClrG
    btnY.BackColor = iClrY
    btnR.BackColor = iClrR
    
    txtG.Text = CStr(iValG)
    txtY.Text = CStr(iValY)
    txtR.Text = CStr(iValR)
End Sub
