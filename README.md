<div align="center">

## Combobox Autofill / Quicken style combobox


</div>

### Description

This class module automatically fills the text of a combo box, using an API call to look up the text from its list.
 
### More Info
 
Dim goAutoFill as New clsComboFill

' In the Change event of the combo box:

Call go_AutoFill.GetListValue(cboBoxName)

' In the KeyUp event of the combo box:

Call go_AutoFill.SupressKeyStroke(cboBox, KeyCode)

Copy this code into a class module called 'clsComboFill.'

Only returns the contents of the combobox's list, or ignores the rest.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mk553](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mk553.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mk553-combobox-autofill-quicken-style-combobox__1-11435/archive/master.zip)

### API Declarations

```
' In the class module:
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
```


### Source Code

```
Option Explicit
' Created by mkeller@hotmail.com - 9/12/2000
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
' Used to hold the keycode supressions
Private m_bSupressKeyCode As Boolean
Private Property Let SupressKeyCode(bValue As Boolean)
  m_bSupressKeyCode = bValue
End Property
Private Property Get SupressKeyCode() As Boolean
  SupressKeyCode = m_bSupressKeyCode
End Property
Public Sub SupressKeyStroke(cboBoxName As ComboBox, KeyCode As Integer)
' This method is called from the KeyDown
' event of a ComboBox.
  ' Let's just assume we only want to supress
  ' backspace and the delete keys.
  If cboBoxName.Text <> "" Then
    Select Case KeyCode
      Case vbKeyDelete
        SupressKeyCode = True
      Case vbKeyBack
        SupressKeyCode = True
    End Select
  End If
End Sub
Public Sub GetListValue(cboBoxName As ComboBox)
' Call this method in the 'Change' event a
' ComboBox.
  Dim lSendMsgContainer As Long, lUnmatchedChars As Long
  Dim sPartialText As String, sTotalText As String
  ' Prevent processing as a result of changes from code
  If m_bSupressKeyCode Then
    m_bSupressKeyCode = False
    Exit Sub
  End If
  With cboBoxName
    ' Lookup list item matching text so far
    sPartialText = .Text
    lSendMsgContainer = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal sPartialText)
    ' If match found, append unmatched characters
    If lSendMsgContainer <> CB_ERR Then
      ' Get full text of matching list item
      sTotalText = .List(lSendMsgContainer)
      ' Compute number of unmatched characters
      lUnmatchedChars = Len(sTotalText) - Len(sPartialText)
      If lUnmatchedChars <> 0 Then
        ' Append unmatched characters to string
        SupressKeyCode = True
        .SelText = Right(sTotalText, lUnmatchedChars)
        ' Select unmatched characters
        .SelStart = Len(sPartialText)
        .SelLength = lUnmatchedChars
      End If
    End If
  End With
End Sub
Private Sub Class_Terminate()
' If there's any kind of err, let's just flush it
' and go about our business. Whoomp, there it
' is!
  Err.Clear
End Sub
```

