VERSION 5.00
Begin VB.UserControl Resizer 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   495
   ToolboxBitmap   =   "Resizer.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Resizer.ctx":0182
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Resizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' if True, also fonts are resized '
Public ResizeFont As Boolean

' if True, form's height/width ratio is preserved '
Public KeepRatio As Boolean

Private Type TControlInfo
       
       ctrl As Control
       Left As Single
       Top As Single
       Width As Single
       Height As Single
       FontSize As Single
       
End Type

Private Type TAllowChanges
  
       AllowChangeTop As Boolean
       AllowChangeLeft As Boolean
       AllowChangeWidth As Boolean
       AllowChangeHeight As Boolean
        
End Type

' this array holds the original position  '
' and size of all controls on parent form '
Dim Controls() As TControlInfo

' a reference to the parent form '
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1

' parent form's size at load time '
Private ParentWidth As Single
Private ParentHeight As Single

' ratio of original height/width '
Private HeightWidthRatio As Single

Private Function CheckForChanges(ByVal TagToUse As String) As TAllowChanges
  
  Dim ChangesToAllow As TAllowChanges
  
  ChangesToAllow.AllowChangeTop = True
  ChangesToAllow.AllowChangeLeft = True
  ChangesToAllow.AllowChangeWidth = True
  ChangesToAllow.AllowChangeHeight = True
    
  If TagToUse <> "" Then
    
    If UCase(Left(TagToUse, 9)) = "MSIRESIZE" Then
      
      ChangesToAllow.AllowChangeTop = False
      ChangesToAllow.AllowChangeLeft = False
      ChangesToAllow.AllowChangeWidth = False
      ChangesToAllow.AllowChangeHeight = False
    
      If Mid(TagToUse, 10, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeLeft = True
        
      End If
      
      If Mid(TagToUse, 11, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeTop = True
        
      End If
      
      If Mid(TagToUse, 12, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeWidth = True
        
      End If
      
      If Mid(TagToUse, 13, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeHeight = True
        
      End If
      
    End If
    
  End If
  
  CheckForChanges = ChangesToAllow
  
End Function

Private Sub ParentForm_Load()

  ' the ParentWidth variable works as a flag '
  ParentWidth = 0
  
  ' save original ratio '
  HeightWidthRatio = ParentForm.Height / ParentForm.Width
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  ResizeFont = PropBag.ReadProperty("ResizeFont", False)
  KeepRatio = PropBag.ReadProperty("KeepRatio", False)
  
  If Ambient.UserMode = False Then
    
    Exit Sub
  
  End If
  
  ' store a reference to the parent form and start receiving events '
  Set ParentForm = Parent
  
End Sub
Private Sub UserControl_Resize()
  
  ' refuse to resize '
  Image1.Move 0, 0
  UserControl.Width = Image1.Width
  UserControl.Height = Image1.Height
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''
' trap the parent form's Resize event      '
' this include the very first resize event '
' that occurs soon after form's load       '
''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParentForm_Resize()
  
  If ParentWidth = 0 Then
    
    Rebuild
  
  Else
    
    Refresh
  
  End If
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' save size and position of all controls on parent form                  '
' you should manually invoke this method each time you add a new control '
' to the form (through Load method of a control array)                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Rebuild()
  
  ' rebuild the internal table
  Dim i As Integer
  Dim ctrl As Control
  
  ' this is necessary for controls that don't support
  ' all properties (e.g. Timer controls)
  On Error Resume Next
    
  If Ambient.UserMode = False Then
    
    Exit Sub
    
  End If
    
  ' save a reference to the parent form, and its initial size
  Set ParentForm = UserControl.Parent
  ParentWidth = ParentForm.ScaleWidth
  ParentHeight = ParentForm.ScaleHeight
    
  ' read the position of all controls on the parent form
  ReDim Controls(ParentForm.Controls.Count - 1) As TControlInfo
    
  For i = 0 To ParentForm.Controls.Count - 1
     
     Set ctrl = ParentForm.Controls(i)
        
     Set Controls(i).ctrl = ctrl
                     
     Controls(i).Left = ctrl.Left
     Controls(i).Top = ctrl.Top
     Controls(i).Width = ctrl.Width
     Controls(i).Height = ctrl.Height
     Controls(i).FontSize = ctrl.Font.Size
     
  Next
  
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update size and position of controls on parent form '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Refresh()
  
  Dim i As Integer
  Dim ctrl As Control
  Dim minFactor As Single
  Dim widthFactor As Single
  Dim heightFactor As Single
  
  Dim Changes As TAllowChanges
      
  On Error Resume Next
  
  ' inhibits recursive calls if KeepRatio = True '
  Static executing As Boolean
  
  If executing Then
    
    Exit Sub
    
  End If
    
  If Ambient.UserMode = False Then
    
    Exit Sub
    
  End If
    
  If ParentForm.WindowState = vbMinimized Then
    
    Exit Sub
    
  End If
  
  If KeepRatio Then
    
    executing = True
    
    ' we must keep original ratio '
    ParentForm.Height = HeightWidthRatio * ParentForm.Width
    executing = False
  
  End If
    
  ' this is necessary for controls that don't support '
  ' all properties (e.g. Timer controls)              '
  On Error Resume Next

  widthFactor = ParentForm.ScaleWidth / ParentWidth
  heightFactor = ParentForm.ScaleHeight / ParentHeight
  
  ' take the lesser of the two '
  If widthFactor < heightFactor Then
    
    minFactor = widthFactor
  
  Else
    
    minFactor = heightFactor
  
  End If
    
  ' this is a regular resize '
  For i = 0 To UBound(Controls)
        
     Changes = CheckForChanges(Controls(i).ctrl.Tag)
     
     ' the change of font must occur *before* the resizing '
     ' to account for companion scrollbar of listbox '
     ' and other similar controls '
     If ResizeFont Then
                       
       Controls(i).ctrl.Font.Size = Controls(i).FontSize * minFactor
                     
     End If
                     
     ' move and resize the controls - we can't use a Move '
     ' method because some controls do not support the change '
     ' of all the four properties (e.g. Height with comboboxes) '
     If Changes.AllowChangeLeft = True Then
                       
       Controls(i).ctrl.Left = Controls(i).Left * widthFactor
                     
     End If
                     
     If Changes.AllowChangeTop = True Then
                       
       Controls(i).ctrl.Top = Controls(i).Top * heightFactor
                     
     End If
                     
     If Changes.AllowChangeWidth = True Then
                       
       Controls(i).ctrl.Width = Controls(i).Width * widthFactor
                     
     End If
                     
     If Changes.AllowChangeHeight = True Then
                       
       Controls(i).ctrl.Height = Controls(i).Height * heightFactor
                     
     End If
     
  Next
  
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  
  Call PropBag.WriteProperty("ResizeFont", ResizeFont, False)
  Call PropBag.WriteProperty("KeepRatio", KeepRatio, False)

End Sub
