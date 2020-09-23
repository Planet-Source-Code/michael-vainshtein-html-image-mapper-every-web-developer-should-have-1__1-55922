VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Mapper by Michael Vainshtein"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveHTML 
      Caption         =   "Save HTML"
      Height          =   285
      Left            =   2850
      TabIndex        =   16
      Top             =   4065
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   4350
      TabIndex        =   15
      Top             =   -45
      Width           =   30
   End
   Begin VB.CommandButton cmdLink 
      Caption         =   "Associate with hyperlink"
      Height          =   285
      Left            =   4440
      TabIndex        =   14
      Top             =   4005
      Width           =   2805
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Preview Page"
      Height          =   285
      Left            =   1477
      TabIndex        =   13
      Top             =   4065
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog DialogBox 
      Left            =   105
      Top             =   4350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Picture..."
      Height          =   285
      Left            =   105
      TabIndex        =   12
      Top             =   4065
      Width           =   1365
   End
   Begin VB.CommandButton cmdDeleteAllObjects 
      Caption         =   "Delete All Areas"
      Height          =   300
      Left            =   5790
      TabIndex        =   11
      ToolTipText     =   "Delete all mapping shapes"
      Top             =   4410
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   4800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDeleteObj 
      Caption         =   "Delete Area"
      Height          =   300
      Left            =   4470
      TabIndex        =   9
      ToolTipText     =   "Delete a mapping shape"
      Top             =   4410
      Width           =   1215
   End
   Begin VB.ListBox lstObjects 
      Height          =   2205
      Left            =   4440
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame frmMode 
      Caption         =   "Mode"
      Height          =   1455
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdFinilizePoly 
         BackColor       =   &H0000FF00&
         Caption         =   "Finilize Poly"
         Enabled         =   0   'False
         Height          =   225
         Left            =   1440
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Finishes the drawind of the polygon"
         Top             =   1065
         Width           =   975
      End
      Begin VB.OptionButton optFreeHand 
         Caption         =   "Polygon"
         Height          =   285
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cretes an area blocked by the polygon"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   285
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Creates an area blocked by a circular shape"
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optRect 
         Caption         =   "Rectangle"
         Height          =   285
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Creates an area blocked by a rectangular shape"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optMove 
         Caption         =   "Pan (move)"
         Height          =   285
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "If you have a big image Pan it to see all of it."
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.PictureBox picOrg 
      Height          =   375
      Left            =   4470
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawMode        =   6  'Mask Pen Not
      DrawWidth       =   2
      ForeColor       =   &H00FF00FF&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      Height          =   240
      Left            =   0
      TabIndex        =   17
      Top             =   -15
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Type POintXY
    X As Integer
    Y As Integer
End Type

Private Type MObjct
    oType As String
    oCoord() As POintXY
    oRadius As Integer
    oDeleted As Boolean
    oLink As String
End Type

Dim FileName, ApPath As String
Dim CurrentObject
Dim XDrag, YDrag, OldY, OldX
Dim Dragging, RectDrag, CircleDrag, FreeHandDrag As Boolean
Dim MapObject() As MObjct
Dim FreeHandIndex As Integer

Private Sub cmdDeleteAllObjects_Click()
    Dim I
    For I = 0 To UBound(MapObject)
        If MapObject(I).oDeleted = True Then GoTo DelOK
    Next
    Exit Sub
DelOK:
    If MsgBox("Delete all mapping objects?", vbYesNo + vbQuestion, "Delete All Mapping Objects") Then
        For I = 0 To UBound(MapObject)
            MapObject(I).oDeleted = True
        Next
        DrawPic 0, 0
        DrawAllObjects
        RefreshList
    End If
End Sub

Private Sub cmdDeleteObj_Click()
    Dim I, C
    If lstObjects.ListIndex <> -1 Then
        For I = 0 To UBound(MapObject)
            If MapObject(I).oDeleted = False Then C = C + 1
            If C = lstObjects.ListIndex + 1 Then MapObject(I).oDeleted = True
        Next

        lstObjects.RemoveItem lstObjects.ListIndex
        DrawPic 0, 0
        DrawAllObjects
    Else: MsgBox "No obejcts selected or no objects exist.", vbExclamation
    End If
End Sub

Private Sub cmdFinilizePoly_Click()
    Dim I
    ReDim Preserve MapObject(CurrentObject).oCoord(FreeHandIndex)
    MapObject(CurrentObject).oCoord(FreeHandIndex).X = MapObject(CurrentObject).oCoord(0).X
    MapObject(CurrentObject).oCoord(FreeHandIndex).Y = MapObject(CurrentObject).oCoord(0).Y
    optMove.Enabled = True: optCircle.Enabled = True: optRect.Enabled = True: cmdDeleteObj.Enabled = True: lstObjects.Enabled = True
    FreeHandDrag = False
    DrawPic 0, 0
    DrawAllObjects OldX, OldY
    RefreshList
    FreeHandIndex = 0
    CurrentObject = CurrentObject + 1
    cmdFinilizePoly.Enabled = False
    Description ""
End Sub

Private Sub cmdFinilizePoly_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Description "Connects the first and last line of the polygon and ends the polygon drawing"
End Sub

Private Sub cmdLink_Click()
    Dim I, C
    If lstObjects.ListIndex <> -1 Then
        For I = 0 To UBound(MapObject)
            If MapObject(I).oDeleted = False Then C = C + 1
            If C = lstObjects.ListIndex + 1 Then
                MapObject(I).oLink = InputBox("What is the link that should be associated with this image area?", "Hyperlink destination", "http://www.")
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdPrev_Click()
    On Error Resume Next
    Kill ApPath & "Test.html"
    Open ApPath & "Test.html" For Binary As #1
        Put #1, , MakeHTML
    Close #1
    ShellFile ApPath & "Test.html"
End Sub

Private Sub cmdLoad_Click()
    On Error Resume Next
    DialogBox.Filter = "All pictures|*.jpg;*.gif;*.jpeg;*.bmp|All Files|*.*"
    DialogBox.ShowOpen
    If DialogBox.FileName <> "" Then
        picOrg.Picture = LoadPicture(DialogBox.FileName)
        FileName = DialogBox.FileName
        cmdDeleteAllObjects_Click
        DrawPic 0, 0
    End If
End Sub

Private Sub cmdSaveHTML_Click()
    On Error Resume Next
    DialogBox.Filter = "HTML Files|*.html;*.htm"
    DialogBox.ShowSave
    If DialogBox.FileName <> "" Then
        Open DialogBox.FileName For Binary As #1
            Put #1, , MakeHTML
        Close #1
    End If
End Sub

Private Sub Form_Load()
    ApPath = App.path & "\"
    If Len(App.path) = 3 Then ApPath = App.path
    DialogBox.FileName = ApPath & "ME.gif"
    FileName = ApPath & "ME.gif"
    picOrg.Picture = LoadPicture(ApPath & "ME.gif")
    Picture1.PaintPicture picOrg.Picture, 0, 0
    CurrentObject = 0
    ReDim MapObject(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Kill "D:\VB\Image Mapper\Test.html"
End Sub

Private Sub Label1_Click()
    Form_Unload (0)
    End
End Sub

Private Sub optCircle_Click()
    Picture1.MousePointer = 0
    Description "Click on the Picture Box to place the center and drag to change the radius"
End Sub

Private Sub optFreeHand_Click()
    Picture1.MousePointer = 0
    Description "Click on the Picture Box to create the polygon's corners. Click Finilize Polygon to finish"
End Sub

Private Sub optMove_Click()
    Picture1.MousePointer = 15
    Description "Click on the picture and drag it to change position"
End Sub

Private Sub optRect_Click()
    Picture1.MousePointer = 0
    Description "Click and drag on the picture box to create a rectangle area"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim I
        If optMove = True Then
            Dragging = True
            XDrag = X
            YDrag = Y
            Picture1.MousePointer = 15
        End If
        If optRect = True Then
            RectDrag = True
            ReDim Preserve MapObject(CurrentObject)
            ReDim Preserve MapObject(CurrentObject).oCoord(1)
            MapObject(CurrentObject).oCoord(0).X = X - OldX
            MapObject(CurrentObject).oCoord(0).Y = Y - OldY
            MapObject(CurrentObject).oDeleted = False
            MapObject(CurrentObject).oType = "rect"
        End If
        If optCircle Then
            ReDim Preserve MapObject(CurrentObject)
            ReDim Preserve MapObject(CurrentObject).oCoord(1)
            MapObject(CurrentObject).oCoord(0).X = X - OldX
            MapObject(CurrentObject).oCoord(0).Y = Y - OldY
            MapObject(CurrentObject).oType = "circle"
            MapObject(CurrentObject).oDeleted = False
            CircleDrag = True
        End If
        If optFreeHand Then
            ReDim Preserve MapObject(CurrentObject)
            ReDim Preserve MapObject(CurrentObject).oCoord(FreeHandIndex)
            MapObject(CurrentObject).oDeleted = False
            MapObject(CurrentObject).oCoord(FreeHandIndex).X = X - OldX
            MapObject(CurrentObject).oCoord(FreeHandIndex).Y = Y - OldY
            MapObject(CurrentObject).oType = "poly"
            FreeHandDrag = True
            FreeHandIndex = FreeHandIndex + 1
            DrawPic 0, 0
            DrawAllObjects OldX, OldY
            For I = 0 To UBound(MapObject(CurrentObject).oCoord) - 1
                Picture1.Line (MapObject(CurrentObject).oCoord(I).X + OldX, MapObject(CurrentObject).oCoord(I).Y + OldY)-(MapObject(CurrentObject).oCoord(I + 1).X + OldX, MapObject(CurrentObject).oCoord(I + 1).Y + OldY)
            Next
            RefreshList
            
            If FreeHandIndex >= 3 Then cmdFinilizePoly.Enabled = True
            optMove.Enabled = False: optCircle.Enabled = False: optRect.Enabled = False: cmdDeleteObj.Enabled = False: lstObjects.Enabled = False
        End If
    ElseIf CircleDrag = True Or RectDrag = True Or FreeHandDrag = True Then
        CircleDrag = False: RectDrag = False: FreeHandDrag = False: FreeHandIndex = 0: cmdFinilizePoly.Enabled = False
        optMove.Enabled = True: optCircle.Enabled = True: optRect.Enabled = True: cmdDeleteObj.Enabled = True: lstObjects.Enabled = True
        MapObject(UBound(MapObject)).oDeleted = True
        CurrentObject = CurrentObject + 1
        RefreshList
        DrawPic 0, 0
        DrawAllObjects
    End If
    Description "Press right mouse button to cancel current operation"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpX, TmpY
    Dim I
    If Dragging = True Then
        DrawPic X, Y
        DrawAllObjects OldX + X - XDrag, OldY + Y - YDrag
    End If
    If RectDrag = True Then
        DrawPic 0, 0
        Picture1.Line (OldX + MapObject(CurrentObject).oCoord(0).X, OldY + MapObject(CurrentObject).oCoord(0).Y)-(OldX + MapObject(CurrentObject).oCoord(0).X, Y)
        Picture1.Line (OldX + MapObject(CurrentObject).oCoord(0).X, OldY + MapObject(CurrentObject).oCoord(0).Y)-(X, OldY + MapObject(CurrentObject).oCoord(0).Y)
        Picture1.Line (X, OldY + MapObject(CurrentObject).oCoord(0).Y)-(X, Y)
        Picture1.Line (OldX + MapObject(CurrentObject).oCoord(0).X, Y)-(X, Y)
        DrawAllObjects OldX, OldY
    End If
    
    If CircleDrag = True Then
        DrawPic 0, 0
        DrawAllObjects OldX, OldY
        Picture1.Circle (MapObject(CurrentObject).oCoord(0).X + OldX, MapObject(CurrentObject).oCoord(0).Y + OldY), Abs(MapObject(CurrentObject).oCoord(0).X + OldX - X)
    End If
    
    If FreeHandDrag = True And FreeHandIndex > 0 Then
        DrawPic 0, 0
        DrawAllObjects OldX, OldY
        For I = 0 To UBound(MapObject(CurrentObject).oCoord) - 1
            Picture1.Line (MapObject(CurrentObject).oCoord(I).X + OldX, MapObject(CurrentObject).oCoord(I).Y + OldY)-(MapObject(CurrentObject).oCoord(I + 1).X + OldX, MapObject(CurrentObject).oCoord(I + 1).Y + OldY)
        Next
        Picture1.Line (MapObject(CurrentObject).oCoord(FreeHandIndex - 1).X + OldX, MapObject(CurrentObject).oCoord(FreeHandIndex - 1).Y + OldY)-(X, Y)
    End If
    StatBar.Panels(StatBar.Panels.Count).Text = "X: " & X - OldX & " Y: " & Y - OldY
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If optMove = True Then
            Dragging = False
            OldX = OldX + X - XDrag
            OldY = OldY + Y - YDrag
            XDrag = 0
            YDrag = 0
            Picture1.MousePointer = 0
        End If
        If optRect = True And RectDrag = True Then
            RectDrag = False
            MapObject(CurrentObject).oCoord(1).X = X - OldX
            MapObject(CurrentObject).oCoord(1).Y = Y - OldY
            CurrentObject = CurrentObject + 1
            RefreshList
            DrawPic 0, 0
            DrawAllObjects OldX, OldY
        End If
        If optCircle = True And CircleDrag = True Then
            CircleDrag = False
            MapObject(CurrentObject).oRadius = Abs(MapObject(CurrentObject).oCoord(0).X - X + OldX)
            DrawPic 0, 0
            DrawAllObjects OldX, OldY
            CurrentObject = CurrentObject + 1
            RefreshList
        End If
    End If
End Sub

Public Function DrawPic(X, Y)
    Picture1.Cls
    Picture1.PaintPicture picOrg.Picture, OldX + (X - XDrag), OldY + (Y - YDrag)
    Picture1.CurrentX = OldX + (X - XDrag) - TextWidth("0,0")
    Picture1.CurrentY = OldY + (Y - YDrag) - TextHeight("0,0")
    Picture1.Print "0,0"
End Function

Public Sub RefreshList()
    Dim I, k, S
    lstObjects.Clear
    For I = 0 To UBound(MapObject)
        S = ""
        If MapObject(I).oDeleted = False Then
            Select Case MapObject(I).oType
                Case "rect": lstObjects.AddItem "Rectangle (" & MapObject(I).oCoord(0).X & ", " & MapObject(I).oCoord(0).Y & "); (" & MapObject(I).oCoord(1).X & "," & MapObject(I).oCoord(1).Y & ")"
                Case "circle": lstObjects.AddItem "Circle (" & MapObject(I).oCoord(0).X & ", " & MapObject(I).oCoord(0).Y & "); Radius=" & MapObject(I).oRadius
                Case "poly":
                    For k = 0 To UBound(MapObject(I).oCoord)
                        S = S & "(" & MapObject(I).oCoord(k).X & ", " & MapObject(I).oCoord(k).Y & ");"
                    Next
                    lstObjects.AddItem "Polygon " & S
            End Select
        End If
    Next
    lstObjects.ListIndex = lstObjects.ListCount - 1
End Sub

Public Sub DrawAllObjects(Optional X = 0, Optional Y = 0)
    Dim I
    For I = 0 To UBound(MapObject)
        If MapObject(I).oDeleted = False Then
            Select Case MapObject(I).oType
                Case "rect":
                    If RectDrag <> True Or I <> UBound(MapObject) Then
                        Picture1.Line (X + MapObject(I).oCoord(0).X, Y + MapObject(I).oCoord(0).Y)-(X + MapObject(I).oCoord(0).X, Y + MapObject(I).oCoord(1).Y)
                        Picture1.Line (X + MapObject(I).oCoord(0).X, Y + MapObject(I).oCoord(0).Y)-(X + MapObject(I).oCoord(1).X, Y + MapObject(I).oCoord(0).Y)
                        Picture1.Line (X + MapObject(I).oCoord(1).X, Y + MapObject(I).oCoord(0).Y)-(X + MapObject(I).oCoord(1).X, Y + MapObject(I).oCoord(1).Y)
                        Picture1.Line (X + MapObject(I).oCoord(0).X, Y + MapObject(I).oCoord(1).Y)-(X + MapObject(I).oCoord(1).X, Y + MapObject(I).oCoord(1).Y)
                    End If
                Case "circle"
                    If CircleDrag <> True Or I <> UBound(MapObject) Then
                        Picture1.Circle (X + MapObject(I).oCoord(0).X, Y + MapObject(I).oCoord(0).Y), MapObject(I).oRadius
                    End If
                Case "poly"
                    If FreeHandDrag <> True Or I <> UBound(MapObject) Then
                        Dim k
                        For k = 0 To UBound(MapObject(I).oCoord) - 1
                            Picture1.Line (MapObject(I).oCoord(k).X + X, MapObject(I).oCoord(k).Y + Y)-(MapObject(I).oCoord(k + 1).X + X, MapObject(I).oCoord(k + 1).Y + Y)
                        Next
                    End If
            End Select
        End If
    Next
End Sub

Public Sub Description(Optional txt = "")
    If txt <> "" Then
        StatBar.Panels(1).AutoSize = sbrContents
        StatBar.Panels(1).Text = txt & ". "
    Else: StatBar.Panels(1).AutoSize = sbrSpring
    End If
    
End Sub

Public Function MakeHTML() As String
    Dim S As String
    Dim I, k
    S = "Right-Click ->View Source.. Copy and paste into your own page<BR>Note that only the ares that you've chosen in the Image Mapper link to their destantion and not the whole rectangular image as usual. (dah, this is the whole purpose of the programme)<BR>" & vbNewLine & "<MAP name=myMap>"
    For I = 0 To UBound(MapObject)
        If MapObject(I).oDeleted = False Then
            Select Case MapObject(I).oType
                Case "circle"
                    S = S & vbNewLine & "     <AREA shape=""circle"" COORDS="""
                    S = S & Str(MapObject(I).oCoord(0).X) & "," & Str(MapObject(I).oCoord(0).Y) & "," & Str(MapObject(I).oRadius)
                Case "rect"
                    S = S & vbNewLine & "     <AREA shape=""rect"" COORDS="""
                    S = S & Str(MapObject(I).oCoord(0).X) & "," & Str(MapObject(I).oCoord(0).Y) & "," & Str(MapObject(I).oCoord(1).X) & "," & Str(MapObject(I).oCoord(1).Y) & """"
                Case "poly"
                    S = S & vbNewLine & "     <AREA shape=""poly"" COORDS="""
                    For k = 0 To UBound(MapObject(I).oCoord)
                        S = S & Str(MapObject(I).oCoord(k).X) & "," & Str(MapObject(I).oCoord(k).Y)
                        If k <> UBound(MapObject(I).oCoord) Then S = S & ","
                    Next
            End Select
            If MapObject(I).oType <> "" Then
                S = S & """ href=""" & MapObject(I).oLink & """>"
            Else: S = "<MAP name=myMap>"
            End If
        End If
    Next
    S = S & vbNewLine & "</MAP>"
    S = S & vbNewLine & vbNewLine & "<IMG src=""" & FileName & """ USEMAP=""#myMap"">"
    MakeHTML = S
End Function


Public Function ShellFile(path As String)
ShellFile = ShellExecute(Me.hwnd, "open", path, "", "", 1)
End Function
