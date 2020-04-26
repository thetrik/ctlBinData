VERSION 5.00
Begin VB.UserControl ctlBinData 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ctlBinData.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlBinData.ctx":0012
End
Attribute VB_Name = "ctlBinData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // ctlBinData.ctl
' // This control allows to store binary files inside executable
' // By The trick 2020
' //

Option Explicit

Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub memset Lib "kernel32" _
                    Alias "RtlFillMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long, _
                    ByVal Fill As Byte)
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function lstrcpyn Lib "kernel32" _
                         Alias "lstrcpynW" ( _
                         ByVal lpString1 As Long, _
                         ByVal lpString2 As Long, _
                         ByVal iMaxLength As Long) As Long
Private Declare Function lstrlen Lib "kernel32" _
                         Alias "lstrlenW" ( _
                         ByVal lpString As Long) As Long
                                
Private Type tFile
    sName   As String
    lSize   As Long
    bData() As Byte
End Type

Private m_tFiles()  As tFile
Private m_lCount    As Long

Public Sub AddFile( _
           ByRef sName As String, _
           ByRef bData() As Byte)
    Dim lIndex  As Long
    
    If IndexByName(sName) <> -1 Then
        Err.Raise 58
    End If
    
    lIndex = m_lCount
    
    If m_lCount Then
        If lIndex > UBound(m_tFiles) Then
            ReDim Preserve m_tFiles(lIndex + 10)
        End If
    Else
        ReDim m_tFiles(9)
    End If
    
    m_tFiles(lIndex).sName = sName
    m_tFiles(lIndex).bData = bData
    m_tFiles(lIndex).lSize = ElementsCount(bData)
    
    m_lCount = m_lCount + 1
    
    PropertyChanged "Content"
        
End Sub

Public Property Get FilesCount() As Long
    FilesCount = m_lCount
End Property

Public Property Get FileSize( _
                    ByRef sName As String) As Long
    Dim lIndex  As Long
    
    lIndex = IndexByName(sName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    FileSize = m_tFiles(lIndex).lSize
          
End Property

Public Property Get FilesList() As String()
    Dim sRet()  As String
    Dim lIndex  As Long
    
    If m_lCount > 0 Then
        
        ReDim sRet(m_lCount - 1)
        
        For lIndex = 0 To m_lCount - 1
            sRet(lIndex) = m_tFiles(lIndex).sName
        Next
        
    End If
    
    FilesList = sRet
    
End Property

Public Property Get File( _
                    ByRef sName As String) As Byte()
    Dim lIndex  As Long
    
    lIndex = IndexByName(sName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    File = m_tFiles(lIndex).bData()
    
End Property

Public Property Let File( _
                    ByRef sName As String, _
                    ByRef bValue() As Byte)
    Dim lIndex  As Long
    
    lIndex = IndexByName(sName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    m_tFiles(lIndex).bData() = bValue
    
    PropertyChanged "Content"
    
End Property

Public Sub RemoveFile( _
           ByRef sName As String)
    Dim lIndex  As Long
    
    lIndex = IndexByName(sName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    Erase m_tFiles(lIndex).bData
    m_tFiles(lIndex).sName = vbNullString
    
    If lIndex < m_lCount - 1 Then
        memcpy m_tFiles(lIndex), m_tFiles(lIndex + 1), (m_lCount - lIndex - 1) * LenB(m_tFiles(lIndex))
        memset m_tFiles(m_lCount - 1), LenB(m_tFiles(lIndex)), 0
    End If
    
    m_lCount = m_lCount - 1
    
    PropertyChanged "Content"
    
End Sub

Public Property Get Exists( _
                    ByRef sName As String) As Boolean
    Exists = IndexByName(sName) <> -1
End Property

Public Sub RenameFile( _
           ByRef sOldName As String, _
           ByRef sNewName As String)
    Dim lIndex  As Long
    
    lIndex = IndexByName(sOldName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    If IndexByName(sNewName) <> -1 Then
        Err.Raise 58
    End If
    
    m_tFiles(lIndex).sName = sNewName
    
    PropertyChanged "Content"
    
End Sub

Public Sub Clear()
    
    Erase m_tFiles
    m_lCount = 0
    
    PropertyChanged "Content"
    
End Sub

Private Function ElementsCount( _
                 ByRef b() As Byte) As Long
    On Error GoTo err_handler
    
    ElementsCount = UBound(b) - LBound(b) + 1
    Exit Function
    
err_handler:
    
    ElementsCount = 0
    
End Function

Private Property Get IndexByName( _
                     ByRef sName As String) As Long
    Dim lIndex  As Long
    
    IndexByName = -1
    
    For lIndex = 0 To m_lCount - 1
        If StrComp(sName, m_tFiles(lIndex).sName, vbTextCompare) = 0 Then
            IndexByName = lIndex
            Exit Property
        End If
    Next
    
End Property

Private Sub UserControl_Paint()
    Dim lWidth  As Long
    
    lWidth = UserControl.TextWidth("UCBinary")
    
    UserControl.Cls
    UserControl.ForeColor = &H8080F0
    UserControl.CurrentX = (UserControl.ScaleWidth - lWidth) / 2
    UserControl.CurrentY = 3

    UserControl.Print "UCBinary"
    UserControl.ForeColor = &H80F080
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), &H505050, B
    
End Sub

Private Sub Deserialize( _
            ByRef bData() As Byte)
    Dim lRemain     As Long
    Dim lCount      As Long
    Dim tFiles()    As tFile
    Dim lIndex      As Long
    Dim lLen        As Long
    Dim lPos        As Long
    
    lRemain = ElementsCount(bData)
    
    If lRemain = 0 Then
    
        m_lCount = 0
        Erase m_tFiles
        Exit Sub
        
    ElseIf lRemain < 4 Then
        Err.Raise 321
    End If
    
    GetMem4 bData(lPos), lCount:   lRemain = lRemain - 4: lPos = lPos + 4
    If lCount < 0 Then
        Err.Raise 321
    End If
    
    If lCount > 0 Then
    
        ReDim tFiles(lCount - 1)
        
        For lIndex = 0 To lCount - 1
            
            If lRemain < 2 Then
                Err.Raise 321
            End If
            
            lLen = lstrlen(VarPtr(bData(lPos)))
            If (lLen + 1) * 2 > lRemain Then
                Err.Raise 321
            End If
            
            tFiles(lIndex).sName = Space$(lLen)
            lstrcpyn StrPtr(tFiles(lIndex).sName), VarPtr(bData(lPos)), lLen + 1
            lRemain = lRemain - (lLen + 1) * 2: lPos = lPos + (lLen + 1) * 2
            
            If lRemain < 4 Then
                Err.Raise 321
            End If
            
            GetMem4 bData(lPos), tFiles(lIndex).lSize: lRemain = lRemain - 4: lPos = lPos + 4
            If lRemain < tFiles(lIndex).lSize Then
                Err.Raise 321
            End If
            
            If tFiles(lIndex).lSize > 0 Then
                ReDim tFiles(lIndex).bData(tFiles(lIndex).lSize - 1)
                memcpy tFiles(lIndex).bData(0), bData(lPos), tFiles(lIndex).lSize
            ElseIf tFiles(lIndex).lSize < 0 Then
                Err.Raise 321
            End If
            
            lPos = lPos + tFiles(lIndex).lSize: lRemain = lRemain - tFiles(lIndex).lSize

        Next
        
    End If
    
    m_lCount = lCount
    m_tFiles = tFiles
    
End Sub
            
Private Function Serialize() As Byte()
    Dim lResultSize As Long
    Dim lIndex      As Long
    Dim bResult()   As Byte
    Dim lPos        As Long
    
    For lIndex = 0 To m_lCount - 1
        lResultSize = lResultSize + ElementsCount(m_tFiles(lIndex).bData) + LenB(m_tFiles(lIndex).sName) + 2
    Next
    
    lResultSize = lResultSize + 4 + m_lCount * 4
    
    ReDim bResult(lResultSize - 1)
    
    GetMem4 m_lCount, bResult(lPos):  lPos = lPos + 4
    
    For lIndex = 0 To m_lCount - 1
        
        memcpy bResult(lPos), ByVal StrPtr(m_tFiles(lIndex).sName), LenB(m_tFiles(lIndex).sName) + 2
        lPos = lPos + LenB(m_tFiles(lIndex).sName) + 2
        GetMem4 m_tFiles(lIndex).lSize, bResult(lPos)
        lPos = lPos + 4
        If m_tFiles(lIndex).lSize Then
            memcpy bResult(lPos), m_tFiles(lIndex).bData(0), m_tFiles(lIndex).lSize
        End If
        lPos = lPos + m_tFiles(lIndex).lSize
               
    Next
    
    Serialize = bResult
    
End Function

Private Sub UserControl_ReadProperties( _
            ByRef cPropBag As PropertyBag)
    Dim bContent()  As Byte
    
    bContent = cPropBag.ReadProperty("Content", bContent())
    
    Deserialize bContent
    
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef cPropBag As PropertyBag)
    cPropBag.WriteProperty "Content", Serialize()
End Sub

Private Sub UserControl_Resize()
    Dim lWidth  As Long
    Dim lHeight As Long
    
    lWidth = UserControl.TextWidth("UCBinary") + 6
    lHeight = UserControl.TextHeight("UCBinary") + 6
    
    UserControl.Width = UserControl.ScaleX(lWidth, vbPixels, vbTwips)
    UserControl.Height = UserControl.ScaleY(lHeight, vbPixels, vbTwips)
    
End Sub

