VERSION 5.00
Begin VB.Form frmCommentCleaner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB6 Comment Cleaner Pro"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VB6CommentCleanerPro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1555
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "VB6CommentCleanerPro.frx":0442
      ToolTipText     =   "Info!"
      Top             =   3100
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cleaning Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   240
      TabIndex        =   3
      Top             =   1550
      Width           =   5535
      Begin VB.OptionButton optMode 
         Caption         =   "Mode 2: Remove Commented Code Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   4815
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Mode 1: Remove All Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clean Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   4980
      Width           =   1695
   End
   Begin VB.TextBox txtDirectory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   5535
   End
   Begin VB.Label lblInfo1 
      Caption         =   "Enter the folder path containing VB6 files (.bas, .cls):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   420
      Width           =   5535
   End
End
Attribute VB_Name = "frmCommentCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================
' Global Settings
' =============================================
Private Enum CleaningMode
    Mode_RemoveAllComments = 0
    Mode_RemoveCommentedCode = 1
End Enum

Private keywords As Variant
Private symbols As String

Private currentMode As CleaningMode
Private processingCanceled As Boolean


' =============================================
' Main Cleaning Routine
' =============================================
Private Sub cmdClean_Click()
    Dim directory As String

    On Error GoTo ErrorHandler
    
    ' Initialize processing
    processingCanceled = False
    Screen.MousePointer = vbHourglass
    cmdClean.Enabled = False
    Refresh
        
    ' Initialize detection parameters
    If IsEmpty(keywords) Then
        'keywords = Split("Dim,If,Then,Else,For,Next,Do,While,Sub,Function,Call,Set,On Error,Select,Case", ",")
        keywords = Split("Dim,If,Then,Else,For,Next,Do,While,Sub,Function,Call,Set,On Error,Select,Case", ",")
        symbols = "=(){}:+-*/^\"
    End If
    
    ' Set cleaning mode
    currentMode = IIf(optMode(0).Value, Mode_RemoveAllComments, Mode_RemoveCommentedCode)
    
    ' Validate path
    directory = Trim(txtDirectory.Text)
    
    If Len(directory) = 0 Or Dir(directory, vbDirectory) = "" Then
        MsgBox "Please enter a valid directory path.", vbExclamation, "Error"
        GoTo CleanExit
    End If
    
    ' Process files with UI responsiveness
    ProcessFiles directory, "*.bas"
    If processingCanceled Then GoTo CleanExit
    ProcessFiles directory, "*.cls"
    
    MsgBox "Process completed successfully.", vbInformation, "Success"
    
CleanExit:
    Screen.MousePointer = vbNormal
    cmdClean.Enabled = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

' =============================================
' File Processing with Multitasking Support
' =============================================
Private Sub ProcessFiles(directory As String, extension As String)
    Dim file As String
    Dim fileCount As Long
    
    On Error GoTo ErrorHandler
    
    file = Dir(directory & "\" & extension)
    
    Do While file <> "" And Not processingCanceled
        ProcessFile directory & "\" & file
        fileCount = fileCount + 1
        
        ' Update UI every 10 files
        If fileCount Mod 10 = 0 Then
            UpdateStatus "Processed " & fileCount & " files..."
            DoEvents  ' Release control to OS
        End If
        
        file = Dir
    Loop
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ProcessFile(filePath As String)
    Dim content As String
    Dim lines() As String
    Dim cleanedLines() As String
    Dim report As String
    Dim i As Long
    Dim lineCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Read file content
    content = ReadFile(filePath)
    lines = Split(content, vbCrLf)
    report = CreateReportHeader(filePath)
    
    ' Process lines with cleanup
    ReDim cleanedLines(UBound(lines))
    lineCount = 0
    
    For i = 0 To UBound(lines)
        ' Allow cancelation check
        If processingCanceled Then Exit Sub
        
        ProcessLine lines(i), report, i + 1
        
        ' Add non-empty lines to cleaned output
        If Len(Trim(lines(i))) > 0 Then
            cleanedLines(lineCount) = lines(i)
            lineCount = lineCount + 1
        End If
        
        ' Maintain responsiveness
        If i Mod 50 = 0 Then DoEvents
    Next
    
    ' Finalize cleaned content
    If lineCount > 0 Then
        ReDim Preserve cleanedLines(lineCount - 1)
    Else
        ReDim cleanedLines(0)
    End If
    
    ' Save results
    WriteFile filePath & ".garbage.log", report
    WriteFile filePath, Join(cleanedLines, vbCrLf)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing: " & filePath & vbCrLf & Err.Description, vbExclamation, "Error"
End Sub

' =============================================
' Core Comment Processing Logic
' Adaptive Line Processing
' =============================================
Private Sub ProcessLine(ByRef line As String, ByRef report As String, lineNumber As Long)
    Dim originalLine As String
    Dim cleanedLine As String
    Dim commentType As String
    
    On Error GoTo ErrorHandler
    
    originalLine = line
    cleanedLine = StripComment(line, commentType)
    
    ' Apply mode-specific cleaning
    Select Case currentMode
        Case Mode_RemoveAllComments
            line = cleanedLine
        Case Mode_RemoveCommentedCode
            If commentType = "CODE" Then
                line = ""
            Else
                line = originalLine  ' Preserve text comments
            End If
    End Select

    ' Generate report
    If originalLine <> line Then
        report = report & _
            "Line " & lineNumber & vbCrLf & _
            "Original: " & originalLine & vbCrLf & _
            "Cleaned:  " & line & vbCrLf & _
            "CommentType: " & commentType & vbCrLf & _
            String(50, "-") & vbCrLf
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' =============================================
' Intelligent Comment Analyzer
' =============================================
Private Function StripComment(line As String, ByRef commentType As String) As String
    Dim inString As Boolean
    Dim commentPos As Long
    Dim i As Long
    Dim char As String
    Dim codePart As String
    Dim commentContent As String
    Dim trimmedcodePart As String
    
    On Error GoTo ErrorHandler
    
    inString = False
    commentPos = 0
    commentType = "TEXT"  ' Default assumption
    
    ' Find comment start position
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        
        If char = """" Then
            inString = Not inString
        ElseIf Not inString And char = "'" Then
            commentPos = i
            Exit For
        End If
    Next
    
    
    ' Apply mode-specific cleaning
    Select Case currentMode
        Case Mode_RemoveAllComments
            If commentPos > 0 Then
                codePart = Left(line, commentPos - 1)
                commentContent = Mid(line, commentPos + 1)
                
                ' Detect code vs text comments
                If IsCodeComment(commentContent) Then
                    commentType = "CODE"
                Else
                    commentType = "TEXT"
                End If
                
                StripComment = codePart
            Else
                StripComment = line
                commentType = "NONE"
            End If
            
        Case Mode_RemoveCommentedCode
            If commentPos > 0 Then
                codePart = Left(line, commentPos - 1)
                commentContent = Mid(line, commentPos + 1)
                
                trimmedcodePart = LCase(Trim(codePart))
                
                ' Empty codePart check
                If Len(trimmedcodePart) = 0 Then
                    ' Detect code vs text comments
                    If IsCodeComment(commentContent) Then
                        commentType = "CODE"
                    Else
                        commentType = "TEXT"
                    End If
                    
                    StripComment = ""
                Else
                    StripComment = codePart
                    commentType = "TEXT"
                End If
            Else
                StripComment = line
                commentType = "NONE"
            End If
    End Select
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

' =============================================
' Advanced Code Detection Logic
' =============================================
Private Function IsCodeComment(content As String) As Boolean
    Dim trimmedContent As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    trimmedContent = LCase(Trim(content))
    
    ' Empty comment check
    If Len(trimmedContent) = 0 Then
        IsCodeComment = False
        Exit Function
    End If
    
    ' Keyword detection
    For i = LBound(keywords) To UBound(keywords)
        If StartsWithCodeWord(trimmedContent, LCase(keywords(i))) Then
            IsCodeComment = True
            Exit Function
        End If
    Next
    
    ' Symbol detection
    For i = 1 To Len(symbols)
        If InStr(trimmedContent, Mid(symbols, i, 1)) > 0 Then
            IsCodeComment = True
            Exit Function
        End If
    Next
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

Private Function StartsWithCodeWord(content As String, word As String) As Boolean
    Dim nextChar As String
    
    On Error GoTo ErrorHandler
    
    If Left(content, Len(word)) = word Then
        ' Check for word boundaries
        If Len(content) > Len(word) Then
            nextChar = Mid(content, Len(word) + 1, 1)
            StartsWithCodeWord = InStr(" " & vbTab & "(", nextChar) > 0
        Else
            StartsWithCodeWord = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

Private Function CreateReportHeader(filePath As String) As String
    
    On Error GoTo ErrorHandler
    
    CreateReportHeader = "Cleaning Mode: " & Choose(currentMode + 1, "Remove All Comments", "Remove Commented Code") & vbCrLf & _
                        "Processed File: " & filePath & vbCrLf & _
                        String(50, "-") & vbCrLf
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

Private Function ReadFile(path As String) As String
    Dim fnum As Integer
    
    On Error GoTo ErrorHandler
    
    fnum = FreeFile
    Open path For Input As #fnum
    ReadFile = Input$(LOF(fnum), fnum)
    Close #fnum
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

Private Sub WriteFile(path As String, content As String)
    Dim fnum As Integer
    
    On Error GoTo ErrorHandler
    
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, content;
    Close #fnum
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' =============================================
' UI Update Utilities
' =============================================
Private Sub UpdateStatus(message As String)
    
    On Error GoTo ErrorHandler
    
    txtInfo.Text = message
    Refresh
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' =============================================
' Cancelation Support
' =============================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    processingCanceled = True
End Sub

