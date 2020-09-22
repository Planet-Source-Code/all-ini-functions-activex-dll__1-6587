<div align="center">

## all Ini Functions activeX\.dll


</div>

### Description

Performs all ini file functions from within a single dll file.
 
### More Info
 
Key name

Section name

New Key Value

You need to add the Inifunctions.dll to the project reference.

then insert the following code in general declaration

eg.

Dim g_IniFunctions As New CInI

then use in your program the following code:

dim RC as variant,R as long,I as long

With g_IniFunctions

SectionId = .SectionGet(Section name)

End With

For R = 0 To UBound(SectionId)

I = InStr(SectionId(R), "=")

Combo1.AddItem Mid(SectionId _(R), I + 1)

Next

dim strReturnWhat as string

strReturnWhat = g_IniFunctions.KeyGet(section name), KeyName)

The other inputs are self explanatory.

ini file name is inside the INIfunctions.dll

save class as Cini.cls and compile as INIFunctions.dll

IniFunction.KeyGet returns the value of a specified key.

IniFunction.SectionGet returns all values in the specified section to a variant

None that I know of, But I have only used it on VB6.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/all-ini-functions-activex-dll__1-6587/archive/master.zip)

### API Declarations

```
Public Declare Function GetPrivateProfileString _ Lib "kernel32" Alias "GetPrivateProfileStringA" _ (ByVal lpApplicationName As String, ByVal _ lpKeyName As Any, ByVal lpDefault As String, _ ByVal lpReturnedString As String, ByVal nSize As _ Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString _ Lib "kernel32" _ alias "WritePrivateProfileStringA" (ByVal _ lpApplicationName As String, ByVal lpKeyName As _ Any, ByVal lpString As Any, ByVal lpFileName As _ String) As Long
Public Declare Function GetPrivateProfileSection _ Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal _ lpAppName As String, ByVal lpReturnedString As _ String, ByVal nSize As Long, ByVal lpFileName As _ String) As Long
```


### Source Code

```
Public m_MstrConfigName As String
Dim m_strKeyname As String
Dim m_strsection As String
Dim m_strKeyValue As String
Dim m_strdefault As String
Private Sub Class_Initialize()
m_MstrConfigName = App.Path & "\ Your Ini file name"
End Sub
Public Property Get KeyName() As String
End Property
Public Property Let KeyName(ByVal strNewValue As String)
End Property
Public Function KeyGet(Optional strSection As String = "N/A", Optional strKeyName = "N/A", Optional strdefault As String = "")
Dim lngRet As Long
'fill in section
If strSection <> "N/A" Then
  m_strsection = strSection
End If
If strKeyName <> "N/A" Then
  m_strKeyname = strKeyName
End If
m_strdefault = strdefault
'get value
m_strKeyValue = Space(255)
lngRet = GetPrivateProfileString(m_strsection, _
                 m_strKeyname, _
                 m_strdefault, _
                 m_strKeyValue, _
                 Len(m_strKeyValue), _
                 m_MstrConfigName)
If lngRet > 0 Then
  m_strKeyValue = Left$(m_strKeyValue, lngRet)
  Else
    m_strKeyValue = vbNullString
End If
 KeyGet = m_strKeyValue
End Function
Public Sub Keysave(Optional strSection As String = "N/A", Optional strKeyName = "N/A", Optional strdefault As String = "")
Dim lngRet As Long
'fill in properties
If strSection <> "N/A" Then
  m_strsection = strSection
End If
If strKeyName <> "N/A" Then
  m_strKeyname = strKeyName
End If
'get value
m_strKeyValue = Space(255)
lngRet = WritePrivateProfileString(m_strsection, _
                 m_strKeyname, _
                 m_strKeyValue, _
                 m_MstrConfigName)
End Sub
Public Function SectionGet(Optional strSection As String = "") As Variant
Dim lngRet As Long
Dim strBuffer As String
If Not strSection = vbNullString Then
  m_strsection = strSection
  End If
If Not m_strsection = vbNullString Then
  strBuffer = Space(2048)
  lngRet = GetPrivateProfileSection(m_strsection, _
                  strBuffer, _
                  Len(strBuffer), _
                  m_MstrConfigName)
 End If
If lngRet > 0 Then
  strBuffer = Left$(strBuffer, lngRet)
  SectionGet = Split(strBuffer, Chr$(0))
  Else
    SectionGet = Array()
End If
End Function
```

