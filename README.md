# Access-VBA-Classes

### SizeAndLocation.cls module

' This class saves and restores your forms size and location between openings.
' To use this class set your Form "Border Style" to "Sizable" and "Min Max Buttons" to "None"
' this code does NOT include support to restore a Form to a Maximized or Minimized state.
'
Option Compare Database
Option Explicit

'  place this at the top of the Forms code module:
Private sl As New SizeAndLocation

'  place this in the Forms Open Event
Private Sub Form_Open(Cancel As Integer)
   Set sl.MyForm = Me.Form
End Sub

'  place this in the Forms Unload Event
Private Sub Form_Unload(Cancel As Integer)
   sl.Save
   Set sl = Nothing
End Sub

