VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'erm, my English is not good
'but I'll do my best :)

'Please visit besoftwaredeveloper.flappie.nl

'Who knows regedit can skip that part
 'Before you start i want to show you something
 'Goto: Start | Run | regedit
 'now click enter; you see it?
 'all the registery of your computer
 'if you want get some time to learn how to use
 'the regedit
 'here we will make sure that everything is OK!
 'now exit the regedit
 'and let's start!
 
'It says delete the value
'from the location we specified
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'Create a new vlaue (String Value)
   'String Value is a value which located on
   'the folders if you saw an icon with red caption
   'picture on it this is a string value
'Where we specifed
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'This opens a values
'We will need it to check if we're not
'setting a new caption to exsiting value
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'This constant says to create a string value
'Win registery API reads it as a string value
'That win registery
Const REG_SZ = 1
'the folder we will locate our value
'this is again a constant of win registery API
Const HKEY_CURRENT_USER = &H80000001
'This is the exact path we will locate our value
'this path is the place of all "startup programs"
'THIS IS NOT A WIN REGISTERY API CONSTANT
Const REGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"
'This says to write a value
'Not to rewrite delete or all other things
'Just write :)
'This is too win registery API const
Const KEY_WRITE = &H20006
'THIS IS A VERY IMPORTANT VARIABLE!!!
'Although it seems quite not important
'this will transfer some data from API to another
Dim Path As Long

Private Sub Command1_Click()
  'If our value is already exist then do not write at all
  'The Path Variable here is taking data about the path
  If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, Path) Then Exit Sub
  'Now the Path Variable knows where is our path and he writes it to there
  'VERY IMPORTANT: Do not replace it with the constants of the path!!!
  'It says to write a new string value with our App title and App Name and path
   'Our path is imprtant else it will start other programs or will not
   'start at all if the value is invaild
  RegSetValueEx Path, App.Title, 0, REG_SZ, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
End Sub

Private Sub Command2_Click()
  'Erm, quite hard to explain this line I am my self didn't got it
  'because if it's exist then exit? but we do need it to be!
  'don't delete that's a non-clear line :s
  If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, Path) Then Exit Sub
  'Delete our string value, locate it with our Variable Path
  'Which I have telled you on the adding all it's story
  'and with our app title
  RegDeleteValue Path, App.Title
End Sub


'That's it! :)
'To make sure all ok you can check it all on regedit on our location
'Hope I helped you! :)

