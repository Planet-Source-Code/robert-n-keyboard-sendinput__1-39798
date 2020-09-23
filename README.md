<div align="center">

## Keyboard \- SendInput


</div>

### Description

The SendInput function synthesizes keystrokes, mouse motions, and button clicks.

Windows NT 4.0 SP3 or later; Windows 98
 
### More Info
 
· nInputs

[in] Specifies the number of structures in the pInputs array.

· pInputs

[in] Pointer to an array of INPUT structures. Each structure represents an event to be inserted into the keyboard or mouse input stream.

· cbSize

[in] Specifies the size, in bytes, of an INPUT structure. If cbSize is not the size of an INPUT structure, the function will fail.

The function returns the number of events that it successfully inserted into the keyboard or mouse input stream. If the function returns zero, the input was already blocked by another thread.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert N\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-n.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-n-keyboard-sendinput__1-39798/archive/master.zip)

### API Declarations

```
Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
```


### Source Code

```
Const VK_H = 72
Const VK_E = 69
Const VK_L = 76
Const VK_O = 79
Const KEYEVENTF_KEYUP = &H2
Const INPUT_MOUSE = 0
Const INPUT_KEYBOARD = 1
Const INPUT_HARDWARE = 2
Private Type MOUSEINPUT
 dx As Long
 dy As Long
 mouseData As Long
 dwFlags As Long
 time As Long
 dwExtraInfo As Long
End Type
Private Type KEYBDINPUT
 wVk As Integer
 wScan As Integer
 dwFlags As Long
 time As Long
 dwExtraInfo As Long
End Type
Private Type HARDWAREINPUT
 uMsg As Long
 wParamL As Integer
 wParamH As Integer
End Type
Private Type GENERALINPUT
 dwType As Long
 xi(0 To 23) As Byte
End Type
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Sub Form_KeyPress(KeyAscii As Integer)
  'Print the key on the form
  Me.Print Chr$(KeyAscii);
End Sub
Private Sub Form_Paint()
  'Clear the form
  Me.Cls
  'call the SendKey-function
  SendKey VK_H
  SendKey VK_E
  SendKey VK_L
  SendKey VK_L
  SendKey VK_O
End Sub
Private Sub SendKey(bKey As Byte)
  Dim GInput(0 To 1) As GENERALINPUT
  Dim KInput As KEYBDINPUT
  KInput.wVk = bKey 'the key we're going to press
  KInput.dwFlags = 0 'press the key
  'copy the structure into the input array's buffer.
  GInput(0).dwType = INPUT_KEYBOARD  ' keyboard input
  CopyMemory GInput(0).xi(0), KInput, Len(KInput)
  'do the same as above, but for releasing the key
  KInput.wVk = bKey ' the key we're going to realease
  KInput.dwFlags = KEYEVENTF_KEYUP ' release the key
  GInput(1).dwType = INPUT_KEYBOARD ' keyboard input
  CopyMemory GInput(1).xi(0), KInput, Len(KInput)
  'send the input now
  Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub
```

