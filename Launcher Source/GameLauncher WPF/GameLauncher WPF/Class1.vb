Imports System.Net
Imports System.IO
Imports System.Diagnostics
Imports Ionic.Zip
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Text.RegularExpressions
Imports System.Collections.ObjectModel

' [ INI File Manager ]
'
' // By Elektro H@cker

#Region " Usage Examples "

'' Set the initialization file path.
'INIFileManager.FilePath = IO.Path.Combine(Application.StartupPath, "Config.ini")

'' Create the initialization file.
'INIFileManager.File.Create()

'' Check that the initialization file exist.
'MsgBox(INIFileManager.File.Exist)

'' Writes a new entire initialization file with the specified text content.
'INIFileManager.File.Write(New List(Of String) From {"[Section Name 1]"})

'' Set an existing value or append it at the end of the initialization file.
'INIFileManager.Key.Set("KeyName1", "Value1")

'' Set an existing value on a specific section or append them at the enf of the initialization file.
'INIFileManager.Key.Set("KeyName2", "Value2", "[Section Name 2]")

'' Gets the value of the specified Key name,
'MsgBox(INIFileManager.Key.Get("KeyName1"))

'' Gets the value of the specified Key name on the specified Section.
'MsgBox(INIFileManager.Key.Get("KeyName2", , "[Section Name 2]"))

'' Gets the value of the specified Key name and returns a default value if the key name is not found.
'MsgBox(INIFileManager.Key.Get("KeyName0", "I'm a default value"))

'' Gets the value of the specified Key name, and assign it to a control property.
'CheckBox1.Checked = CType(INIFileManager.Key.Get("KeyName1"), Boolean)

'' Checks whether a Key exists.
'MsgBox(INIFileManager.Key.Exist("KeyName1"))

'' Checks whether a Key exists on a specific section.
'MsgBox(INIFileManager.Key.Exist("KeyName2", "[First Section]"))

'' Remove a key name.
'INIFileManager.Key.Remove("KeyName1")

'' Remove a key name on the specified Section.
'INIFileManager.Key.Remove("KeyName2", "[Section Name 2]")

'' Add a new section.
'INIFileManager.Section.Add("[Section Name 3]")

'' Get the contents of a specific section.
'MsgBox(String.Join(Environment.NewLine, INIFileManager.Section.Get("[Section Name 1]")))

'' Remove an existing section.
'INIFileManager.Section.Remove("[Section Name 2]")

'' Checks that the initialization file contains at least one section.
'MsgBox(INIFileManager.Section.Has())

'' Sort the initialization file (And remove empty lines).
'INIFileManager.File.Sort(True)

'' Gets the initialization file section names.
'MsgBox(String.Join(", ", INIFileManager.Section.GetNames()))

'' Gets the initialization file content.
'MsgBox(String.Join(Environment.NewLine, INIFileManager.File.Get()))

'' Delete the initialization file from disk.
'INIFileManager.File.Delete()

#End Region

#Region " INI File Manager "

Public Class INIFileManager

#Region " Members "

#Region " Properties "

    ''' <summary>
    ''' Indicates the initialization file path.
    ''' </summary>
    Public Shared Property FilePath As String =
        IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), Process.GetCurrentProcess().ProcessName & ".ini")

#End Region

#Region " Variables "

    ''' <summary>
    ''' Stores the initialization file content.
    ''' </summary>
    Private Shared Content As New List(Of String)

    ''' <summary>
    ''' Stores the INI section names.
    ''' </summary>
    Private Shared SectionNames As String() = {String.Empty}

    ''' <summary>
    ''' Indicates the start element index of a section name.
    ''' </summary>
    Private Shared SectionStartIndex As Integer = -1

    ''' <summary>
    ''' Indicates the end element index of a section name.
    ''' </summary>
    Private Shared SectionEndIndex As Integer = -1

    ''' <summary>
    ''' Stores a single sorted section block with their keys and values.
    ''' </summary>
    Private Shared SortedSection As New List(Of String)

    ''' <summary>
    ''' Stores all the sorted section blocks with their keys and values.
    ''' </summary>
    Private Shared SortedSections As New List(Of String)

    ''' <summary>
    ''' Indicates the INI element index that contains the Key and value.
    ''' </summary>
    Private Shared KeyIndex As Integer = -1

    ''' <summary>
    ''' Indicates the culture to compare the strings.
    ''' </summary>
    Private Shared ReadOnly CompareMode As StringComparison = StringComparison.InvariantCultureIgnoreCase

#End Region

#Region " Exceptions "

    ''' <summary>
    ''' Exception is thrown when a section name parameter has invalid format.
    ''' </summary>
    Private Class SectionNameInvalidFormatException
        Inherits Exception

        Public Sub New()
            MyBase.New("Section name parameter has invalid format." &
                       Environment.NewLine &
                       "The rigth syntax is: [SectionName]")
        End Sub

        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        Public Sub New(message As String, inner As Exception)
            MyBase.New(message, inner)
        End Sub

    End Class

#End Region

#End Region

#Region " Methods "

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Private Shadows Sub ReferenceEquals()
    End Sub

    <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
    Private Shadows Sub Equals()
    End Sub

    Public Class [File]

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub ReferenceEquals()
        End Sub

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub Equals()
        End Sub

        ''' <summary>
        ''' Checks whether the initialization file exist.
        ''' </summary>
        ''' <returns>True if initialization file exist, otherwise False.</returns>
        Public Shared Function Exist() As Boolean
            Return IO.File.Exists(FilePath)
        End Function

        ''' <summary>
        ''' Creates the initialization file.
        ''' If the file already exist it would be replaced.
        ''' </summary>
        ''' <param name="Encoding">The Text encoding to write the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Create(Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            Try
                IO.File.WriteAllText(FilePath,
                                     String.Empty,
                                     If(Encoding Is Nothing, System.Text.Encoding.Default, Encoding))
            Catch ex As Exception
                Throw
                Return False

            End Try

            Return True

        End Function

        ''' <summary>
        ''' Deletes the initialization file.
        ''' </summary>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Delete() As Boolean

            If Not [File].Exist Then Return False

            Try
                IO.File.Delete(FilePath)
            Catch ex As Exception
                Throw
                Return False

            End Try

            Content = Nothing

            Return True

        End Function

        ''' <summary>
        ''' Returns the initialization file content.
        ''' </summary>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        Public Shared Function [Get](Optional ByVal Encoding As System.Text.Encoding = Nothing) As List(Of String)

            Content = IO.File.ReadAllLines(FilePath,
                                           If(Encoding Is Nothing, System.Text.Encoding.Default, Encoding)).ToList()

            Return Content

        End Function

        ''' <summary>
        ''' Sort the initialization file content by the Key names.
        ''' If the initialization file contains sections then the sections are sorted by their names also.
        ''' </summary>
        ''' <param name="RemoveEmptyLines">Remove empty lines.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Sort(Optional ByVal RemoveEmptyLines As Boolean = False,
                                    Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then Return False

            [File].[Get](Encoding)

            Select Case Section.Has(Encoding)

                Case True ' initialization file contains at least one Section.

                    SortedSection.Clear()
                    SortedSections.Clear()

                    Section.GetNames(Encoding) ' Get the (sorted) section names

                    For Each name As String In SectionNames

                        SortedSection = Section.[Get](name, Encoding) ' Get the single section lines.

                        If RemoveEmptyLines Then ' Remove empty lines.
                            SortedSection = SortedSection.Where(Function(line) _
                                                                Not String.IsNullOrEmpty(line) AndAlso
                                                                Not String.IsNullOrWhiteSpace(line)).ToList
                        End If

                        SortedSection.Sort() ' Sort the single section keys.

                        SortedSections.Add(name) ' Add the section name to the sorted sections list.
                        SortedSections.AddRange(SortedSection) ' Add the single section to the sorted sections list.

                    Next name

                    Content = SortedSections

                Case False ' initialization file doesn't contains any Section.
                    Content.Sort()

                    If RemoveEmptyLines Then
                        Content = Content.Where(Function(line) _
                                                        Not String.IsNullOrEmpty(line) AndAlso
                                                        Not String.IsNullOrWhiteSpace(line)).ToList
                    End If

            End Select ' Section.Has()

            ' Save changes.
            Return [File].Write(Content, Encoding)

        End Function

        ''' <summary>
        ''' Writes a new initialization file with the specified text content..
        ''' </summary>
        ''' <param name="Content">Indicates the text content to write in the initialization file.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Write(ByVal Content As List(Of String),
                                     Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            Try
                IO.File.WriteAllLines(FilePath,
                                      Content,
                                      If(Encoding Is Nothing, System.Text.Encoding.Default, Encoding))
            Catch ex As Exception
                Throw
                Return False

            End Try

            Return True

        End Function

    End Class

    Public Class [Key]

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub ReferenceEquals()
        End Sub

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub Equals()
        End Sub

        ''' <summary>
        ''' Return a value indicating whether a key name exist or not.
        ''' </summary>
        ''' <param name="KeyName">Indicates the key name that contains the value to modify.</param>
        ''' <param name="SectionName">Indicates the Section name where to find the key name.</param>
        ''' <param name="Encoding">The Text encoding to write the initialization file.</param>
        ''' <returns>True if the key name exist, otherwise False.</returns>
        Public Shared Function Exist(ByVal KeyName As String,
                                     Optional ByVal SectionName As String = Nothing,
                                     Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then Return False

            [File].[Get](Encoding)

            [Key].GetIndex(KeyName, SectionName)

            Select Case SectionName Is Nothing

                Case True
                    Return Convert.ToBoolean(Not KeyIndex)

                Case Else
                    Return Convert.ToBoolean(Not (KeyIndex + SectionStartIndex))

            End Select

        End Function

        ''' <summary>
        ''' Set the value of an existing key name.
        ''' 
        ''' If the initialization file doesn't exists, or else the Key doesn't exist,
        ''' or else the Section parameter is not specified and the key name doesn't exist;
        ''' then the 'key=value' is appended to the end of the initialization file.
        ''' 
        ''' if the specified Section name exist but the Key name doesn't exist,
        ''' then the 'key=value' is appended to the end of the Section.
        ''' 
        ''' </summary>
        ''' <param name="KeyName">Indicates the key name that contains the value to modify.</param>
        ''' <param name="Value">Indicates the new value.</param>
        ''' <param name="SectionName">Indicates the Section name where to find the key name.</param>
        ''' <param name="Encoding">The Text encoding to write the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function [Set](ByVal KeyName As String,
                                     ByVal Value As String,
                                     Optional ByVal SectionName As String = Nothing,
                                     Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then [File].Create()

            [File].[Get](Encoding)

            [Key].GetIndex(KeyName, SectionName)

            ' If KeyName is not found and indicated Section is found, then...
            If KeyIndex = -1 AndAlso SectionEndIndex <> -1 Then

                ' If section EndIndex is the last line of file, then...
                If SectionEndIndex = Content.Count Then

                    Content(Content.Count - 1) = Content(Content.Count - 1) &
                                                         Environment.NewLine &
                                                         String.Format("{0}={1}", KeyName, Value)

                Else ' If not section EndIndex is the last line of file, then...

                    Content(SectionEndIndex) = String.Format("{0}={1}", KeyName, Value) &
                                                    Environment.NewLine &
                                                    Content(SectionEndIndex)
                End If

                ' If KeyName is found then...
            ElseIf KeyIndex <> -1 Then
                Content(KeyIndex) = String.Format("{0}={1}", KeyName, Value)

                ' If KeyName is not found and Section parameter is passed. then...
            ElseIf KeyIndex = -1 AndAlso SectionName IsNot Nothing Then
                Content.Add(SectionName)
                Content.Add(String.Format("{0}={1}", KeyName, Value))

                ' If KeyName is not found, then...
            ElseIf KeyIndex = -1 Then
                Content.Add(String.Format("{0}={1}", KeyName, Value))

            End If

            ' Save changes.
            Return [File].Write(Content, Encoding)

        End Function

        ''' <summary>
        ''' Get the value of an existing key name.
        ''' If the initialization file or else the Key doesn't exist then a 'Nothing' object is returned. 
        ''' </summary>
        ''' <param name="KeyName">Indicates the key name to retrieve their value.</param>
        ''' <param name="DefaultValue">Indicates a default value to return if the key name is not found.</param>
        ''' <param name="SectionName">Indicates the Section name where to find the key name.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        Public Shared Function [Get](ByVal KeyName As String,
                                     Optional ByVal DefaultValue As Object = Nothing,
                                     Optional ByVal SectionName As String = Nothing,
                                     Optional ByVal Encoding As System.Text.Encoding = Nothing) As Object

            If Not [File].Exist() Then Return DefaultValue

            [File].[Get](Encoding)

            [Key].GetIndex(KeyName, SectionName)

            Select Case KeyIndex

                Case Is <> -1 ' KeyName found.
                    Return Content(KeyIndex).Substring(Content(KeyIndex).IndexOf("=") + 1)

                Case Else ' KeyName not found.
                    Return DefaultValue

            End Select

        End Function

        ''' <summary>
        ''' Returns the initialization file line index of the key name.
        ''' </summary>
        ''' <param name="KeyName">Indicates the Key name to retrieve their value.</param>
        ''' <param name="SectionName">Indicates the Section name where to find the key name.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        Private Shared Sub GetIndex(ByVal KeyName As String,
                                    Optional ByVal SectionName As String = Nothing,
                                    Optional ByVal Encoding As System.Text.Encoding = Nothing)

            If Content Is Nothing Then [File].Get(Encoding)

            ' Reset the INI index elements to negative values.
            KeyIndex = -1
            SectionStartIndex = -1
            SectionEndIndex = -1

            If SectionName IsNot Nothing AndAlso Not SectionName Like "[[]?*[]]" Then
                Throw New SectionNameInvalidFormatException
                Exit Sub
            End If

            ' Locate the KeyName and set their element index.
            ' If the KeyName is not found then the value is set to "-1" to return an specified default value.
            Select Case String.IsNullOrEmpty(SectionName)

                Case True ' Any SectionName parameter is specified.

                    KeyIndex = Content.FindIndex(Function(line) line.StartsWith(String.Format("{0}=", KeyName),
                                                                              StringComparison.InvariantCultureIgnoreCase))

                Case False ' SectionName parameter is specified.

                    Select Case Section.Has(Encoding)

                        Case True ' INI contains at least one Section.

                            SectionStartIndex = Content.FindIndex(Function(line) line.Trim.Equals(SectionName.Trim, CompareMode))
                            If SectionStartIndex = -1 Then ' Section doesn't exist.
                                Exit Sub
                            End If

                            SectionEndIndex = Content.FindIndex(SectionStartIndex + 1, Function(line) line.Trim Like "[[]?*[]]")
                            If SectionEndIndex = -1 Then
                                ' This fixes the value if the section is at the end of file.
                                SectionEndIndex = Content.Count
                            End If

                            KeyIndex = Content.FindIndex(SectionStartIndex, SectionEndIndex - SectionStartIndex,
                                                                  Function(line) line.StartsWith(String.Format("{0}=", KeyName),
                                                                                      StringComparison.InvariantCultureIgnoreCase))

                        Case False ' INI doesn't contains Sections.
                            GetIndex(KeyName, , Encoding)

                    End Select ' Section.Has()

            End Select ' String.IsNullOrEmpty(SectionName)

        End Sub

        ''' <summary>
        ''' Remove an existing key name.
        ''' </summary>
        ''' <param name="KeyName">Indicates the key name to retrieve their value.</param>
        ''' <param name="SectionName">Indicates the Section name where to find the key name.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Remove(ByVal KeyName As String,
                                      Optional ByVal SectionName As String = Nothing,
                                      Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then Return False

            [File].[Get](Encoding)

            [Key].GetIndex(KeyName, SectionName)

            Select Case KeyIndex

                Case Is <> -1 ' Key found.

                    ' Remove the element containing the key name.
                    Content.RemoveAt(KeyIndex)

                    ' Save changes.
                    Return [File].Write(Content, Encoding)

                Case Else ' KeyName not found.
                    Return False

            End Select

        End Function

    End Class

    Public Class Section

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub ReferenceEquals()
        End Sub

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Private Shadows Sub Equals()
        End Sub

        ''' <summary>
        ''' Adds a new section at bottom of the initialization file.
        ''' </summary>
        ''' <param name="SectionName">Indicates the Section name to add.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Add(Optional ByVal SectionName As String = Nothing,
                                   Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then [File].Create()

            If Not SectionName Like "[[]?*[]]" Then
                Throw New SectionNameInvalidFormatException
                Exit Function
            End If

            [File].[Get](Encoding)

            Select Case Section.GetNames(Encoding).Where(Function(line) line.Trim.Equals(SectionName.Trim, CompareMode)).Any

                Case False ' Any of the existing Section names is equal to given section name.

                    ' Add the new section name.
                    Content.Add(SectionName)

                    ' Save changes.
                    Return [File].Write(Content, Encoding)

                Case Else ' An existing Section name is equal to given section name.
                    Return False

            End Select

        End Function

        ''' <summary>
        ''' Returns all the keys and values of an existing Section Name.
        ''' </summary>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <param name="SectionName">Indicates the section name where to retrieve their keynames and values.</param>
        Public Shared Function [Get](ByVal SectionName As String,
                                     Optional ByVal Encoding As System.Text.Encoding = Nothing) As List(Of String)

            If Content Is Nothing Then [File].Get(Encoding)

            SectionStartIndex = Content.FindIndex(Function(line) line.Trim.Equals(SectionName.Trim, CompareMode))

            SectionEndIndex = Content.FindIndex(SectionStartIndex + 1, Function(line) line.Trim Like "[[]?*[]]")

            If SectionEndIndex = -1 Then
                SectionEndIndex = Content.Count ' This fixes the value if the section is at the end of file.
            End If

            Return Content.GetRange(SectionStartIndex, SectionEndIndex - SectionStartIndex).Skip(1).ToList

        End Function

        ''' <summary>
        ''' Returns all the section names of the initialization file.
        ''' </summary>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        Public Shared Function GetNames(Optional ByVal Encoding As System.Text.Encoding = Nothing) As String()

            If Content Is Nothing Then [File].Get(Encoding)

            ' Get the Section names.
            SectionNames = (From line In Content Where line.Trim Like "[[]?*[]]").ToArray

            ' Sort the Section names.
            If SectionNames.Count <> 0 Then Array.Sort(SectionNames)

            ' Return the Section names.
            Return SectionNames

        End Function

        ''' <summary>
        ''' Gets a value indicating whether the initialization file contains at least one Section.
        ''' </summary>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the INI contains at least one section, otherwise False.</returns>
        Public Shared Function Has(Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Content Is Nothing Then [File].Get(Encoding)

            Return (From line In Content Where line.Trim Like "[[]?*[]]").Any()

        End Function

        ''' <summary>
        ''' Removes an existing section with all of it's keys and values.
        ''' </summary>
        ''' <param name="SectionName">Indicates the Section name to remove with all of it's key/values.</param>
        ''' <param name="Encoding">The Text encoding to read the initialization file.</param>
        ''' <returns>True if the operation success, otherwise False.</returns>
        Public Shared Function Remove(Optional ByVal SectionName As String = Nothing,
                                      Optional ByVal Encoding As System.Text.Encoding = Nothing) As Boolean

            If Not [File].Exist() Then Return False

            If Not SectionName Like "[[]?*[]]" Then
                Throw New SectionNameInvalidFormatException
                Exit Function
            End If

            [File].[Get](Encoding)

            Select Case [Section].GetNames(Encoding).Where(Function(line) line.Trim.Equals(SectionName.Trim, CompareMode)).Any

                Case True ' An existing Section name is equal to given section name.

                    ' Get the section StartIndex and EndIndex.
                    [Get](SectionName)

                    ' Remove the section range index.
                    Content.RemoveRange(SectionStartIndex, SectionEndIndex - SectionStartIndex)

                    ' Save changes.
                    Return [File].Write(Content, Encoding)

                Case Else ' Any of the existing Section names is equal to given section name.
                    Return False

            End Select

        End Function

    End Class

#End Region

End Class

#End Region

Public Class INIUtils

    Public Shared Sub Set_Value(File As String, ValueName As String, Value As String)

        Try
            ' Create a new INI File with "Key=Value""
            If Not System.IO.File.Exists(File) Then

                System.IO.File.WriteAllText(File, ValueName & "=" & Value)

                ' Search line by line in the INI file for the "Key"
                Return
            Else

                Dim Line_Number As Int64 = 0
                Dim strArray As String() = System.IO.File.ReadAllLines(File)

                For Each line_loopVariable As String In strArray
                    Dim line As String = line_loopVariable
                    If line.StartsWith(ValueName & "=") Then
                        strArray(Line_Number) = ValueName & "=" & Value
                        System.IO.File.WriteAllLines(File, strArray)
                        ' Replace "value"
                        Return
                    End If
                    Line_Number += 1
                Next

                ' Key don't exist, then create the new "Key=Value"

                System.IO.File.AppendAllText(File, Environment.NewLine + ValueName & "=" & Value)

            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Public Shared Function Load_Value(ValueName As String, Optional File As String = "", Optional Text As String = "") As String

        If String.IsNullOrEmpty(File) AndAlso String.IsNullOrEmpty(Text) Then
            Throw New Exception("At least one of the two nullable parameters has to be filled.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) AndAlso Not String.IsNullOrEmpty(Text) Then
            Throw New Exception("File and text cannot be in the same function called.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) Then
            If Not System.IO.File.Exists(File) Then

                Throw New Exception(File & " not found.")

                ' INI File not found.
                Return Nothing

            Else

                Dim returnvalue As String = ""

                For Each line_loopVariable As String In System.IO.File.ReadAllLines(File)
                    If line_loopVariable.StartsWith(ValueName & "=") Then
                        returnvalue = line_loopVariable.Split("="c).Last()
                    End If
                Next

                If String.IsNullOrEmpty(returnvalue) Then

                    Throw New Exception("Key: " & """" & ValueName & """" & " not found. On " & File)

                    ' Key not found.
                    Return Nothing

                Else
                    Return returnvalue
                End If
            End If
        ElseIf Not String.IsNullOrEmpty(Text) Then

            Dim returnvalue As String = ""

            For Each line_loopVariable As String In Text.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                If line_loopVariable.StartsWith(ValueName & "=") Then
                    returnvalue = line_loopVariable.Split("="c).Last()
                End If
            Next

            If String.IsNullOrEmpty(returnvalue) Then

                Throw New Exception("Key: " & """" & ValueName & """" & " not found. On " & Text)
                ' Key not found.

                Return Nothing

            Else
                Return returnvalue
            End If
        End If

        Return Nothing

    End Function


    Public Shared Sub Delete_Value(File As String, ValueName As String)

        If Not System.IO.File.Exists(File) Then

            Throw New Exception(File & " not found.")

            ' INI File not found.
            Exit Sub

        Else

            Try
                Dim Line_Number As Int64 = 0
                Dim strArray As String() = System.IO.File.ReadAllLines(File)

                For Each line_loopVariable As String In strArray
                    Dim line As String = line_loopVariable
                    If line.StartsWith(ValueName & "=") Then
                        strArray(Line_Number) = Nothing
                        Exit For
                    End If
                    Line_Number += 1
                Next

                Array.Copy(strArray, Line_Number + 1, strArray, Line_Number, strArray.Length - Line_Number)
                Array.Resize(strArray, strArray.Length)

                System.IO.File.WriteAllText(File, String.Join(Environment.NewLine, strArray))

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If

    End Sub


    Public Shared Sub Sort_Values(File As String)

        If Not System.IO.File.Exists(File) Then

            Throw New Exception(File & " not found.")

            ' INI File not found.
            Exit Sub

        Else

            Try
                Dim Line_Number As Int64 = 0
                Dim strArray As String() = System.IO.File.ReadAllLines(File)
                Dim TempList As New List(Of String)()

                For Each line As String In strArray
                    If Not String.IsNullOrEmpty(line) Then
                        TempList.Add(strArray(Line_Number))
                    End If
                    Line_Number += 1
                Next

                TempList.Sort()

                System.IO.File.WriteAllLines(File, TempList.ToArray())
            Catch ex As Exception
                Throw New Exception(ex.Message)

            End Try
        End If

    End Sub

    Public Shared Sub Create_Group(File As String, Name As String, Keys As Dictionary(Of String, String))

        If Not System.IO.File.Exists(File) Then

            'Throw New Exception(File & " not found.")

            ' INI File not found.
            'Exit Sub

        End If

        'Else
        Dim strToPut As String = Name & " {" & Environment.NewLine
        For Each item As KeyValuePair(Of String, String) In Keys
            strToPut += item.Key & "=" & item.Value & Environment.NewLine
        Next
        strToPut += "}"
        System.IO.File.AppendAllText(File, strToPut)
        'End If

    End Sub

    Public Shared Function Read_Group(Name As String, Optional File As String = "", Optional FromText As String = "") As String

        If String.IsNullOrEmpty(File) AndAlso String.IsNullOrEmpty(FromText) Then
            Throw New Exception("At least one of the two nullable parameters has to be filled.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) AndAlso Not String.IsNullOrEmpty(FromText) Then
            Throw New Exception("File and text cannot be in the same function called.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) Then

            If Not System.IO.File.Exists(File) Then

                Throw New Exception(File & " not found.")

                ' INI File not found.
                Return Nothing

            Else

                Dim AllText As String = System.IO.File.ReadAllText(File)

                Dim enclosedText As MatchCollection = Regex.Matches(AllText, "(?<=[\{]).+(?=[\}])", RegexOptions.Singleline)

                Dim Content As String = ""

                For Each mth As Match In enclosedText
                    Content += mth.Value
                Next

                'AllText.Substring(AllText.IndexOf(Name + " {"), AllText.IndexOf("}", AllText.IndexOf(Name + " {")) - AllText.IndexOf(Name + " {"));

                If String.IsNullOrEmpty(Content) Then

                    Throw New Exception("Key: " & """" & Name & """" & " not found.")

                    ' Key not found.
                    Return Nothing

                Else
                    Return Content
                End If

            End If
        ElseIf Not String.IsNullOrEmpty(FromText) Then

            Dim enclosedText As MatchCollection = Regex.Matches(FromText, "(?<=[\{]).+(?=[\}])", RegexOptions.Singleline)

            Dim Content As String = ""

            For Each mth As Match In enclosedText
                Content += mth.Value
            Next
            If String.IsNullOrEmpty(Content) Then

                Throw New Exception("Key: " & """" & Name & """" & " not found.")

                ' Key not found.
                Return Nothing

            Else
                Return Content
            End If
        End If

        Return Nothing

    End Function

    Public Shared Sub Edit_Group(File As String, Name As String, KeysToSet As Dictionary(Of String, String))

        If KeysToSet.Count = 0 Then
            Throw New Exception("Las claves a revisar no pueden ser un valor nulo.")
        End If

        'Get all from a file an compare them

        Dim KeyToCompare As Dictionary(Of String, String) = Extract_Keys("", Read_Group(Name, File))

        Dim FinalKeyDic As Dictionary(Of String, String) = New Dictionary(Of String, String)

        Dim DifferentKeys As Dictionary(Of String, String) = New Dictionary(Of String, String)

        For Each item As KeyValuePair(Of String, String) In KeysToSet

            If KeyToCompare.ContainsKey(item.Key) Or Not KeyToCompare.ContainsKey(item.Key) Then
                FinalKeyDic.Add(item.Key, item.Value) 'Edit the key or add it, it doesn't matter
            End If

        Next

        'But if the keytoset don't have ther key from the default we have to add it.

        DifferentKeys = KeyToCompare.Except(KeysToSet)

        For Each item As KeyValuePair(Of String, String) In DifferentKeys
            FinalKeyDic.Add(item.Key, item.Value)
        Next

        'And finally we have to set it into a file

        Delete_Group(File, Name)

        Create_Group(File, Name, FinalKeyDic)

    End Sub

    Public Shared Sub Delete_Group(File As String, Name As String)

        'Find the exact part where the line is

        Dim FileContent As String = IO.File.ReadAllText(File)

        FileContent = Regex.Replace(FileContent, Name & " {.*}", "", RegexOptions.Singleline)

        IO.File.WriteAllText(File, FileContent)

    End Sub

    Public Shared Function Group_Exists(File As String, Name As String) As Boolean
        If Not Directory.Exists(Path.GetDirectoryName(File)) Then
            Return False
        ElseIf Not IO.File.Exists(File) Then
            Return False
        ElseIf Not System.IO.File.ReadAllText(File).Contains(Name & " {") Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function Extract_Keys(Optional File As String = "", Optional Text As String = "") As Dictionary(Of String, String)

        If String.IsNullOrEmpty(File) AndAlso String.IsNullOrEmpty(Text) Then
            Throw New Exception("At least one of the two nullable parameters has to be filled.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) AndAlso Not String.IsNullOrEmpty(Text) Then
            Throw New Exception("File and text cannot be in the same function called.")
            Return Nothing
        End If

        If Not String.IsNullOrEmpty(File) Then
            If Not System.IO.File.Exists(File) Then

                Throw New Exception(File & " not found.")

                ' INI File not found.
                Return Nothing

            Else

                Dim returnvalue As New Dictionary(Of String, String)()

                Dim FileText As String = System.IO.File.ReadAllText(File)

                For Each line_loopVariable As String In FileText.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    Dim key As String = line_loopVariable.Split("="c).First()
                    Dim value As String = line_loopVariable.Split("="c).Last()
                    returnvalue.Add(key, value)
                Next

                If returnvalue Is Nothing Then

                    Throw New Exception("No se encontro ninguna key dentro de este archivo.")

                    ' Key not found.
                    Return Nothing

                Else
                    Return returnvalue
                End If

            End If
        ElseIf Not String.IsNullOrEmpty(Text) Then

            Dim returnvalue As New Dictionary(Of String, String)()

            For Each line_loopVariable As String In Text.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Dim key As String = line_loopVariable.Split("="c).First()
                Dim value As String = line_loopVariable.Split("="c).Last()
                returnvalue.Add(key, value)
            Next

            If returnvalue Is Nothing Then

                Throw New Exception("No se encontro ninguna key dentro de este texto.")

                ' Key not found.
                Return Nothing

            Else
                Return returnvalue
            End If
        End If

        Return Nothing

    End Function

End Class

Public Class TreeHelper

    Public Shared Function FindVisualChildByName(Of T As FrameworkElement)(parent As DependencyObject, name As String) As T
        Dim child As T = Nothing
        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim ch = VisualTreeHelper.GetChild(parent, i)
            child = TryCast(ch, T)
            If child IsNot Nothing AndAlso child.Name = name Then
                Exit For
            Else
                child = FindVisualChildByName(Of T)(ch, name)
            End If

            If child IsNot Nothing Then
                Exit For
            End If
        Next
        Return child
    End Function

End Class

Public Class ZipManager

    Public Shared Sub Extract(ZipAExtraer As String, PathToDownloadExtraccion As String)
        Try

            Using zip1 As ZipFile = ZipFile.Read(ZipAExtraer)
                Dim e As ZipEntry
                For Each e In zip1
                    e.Extract(PathToDownloadExtraccion, ExtractExistingFileAction.OverwriteSilently)
                Next
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Shared Sub Compress(NombrePathToDownload, NombreGuardar)
        Using zip As ZipFile = New ZipFile()
            zip.AddDirectory(NombrePathToDownload)
            zip.Save(NombreGuardar)
        End Using
    End Sub

End Class

Public Class Updater

    Public Shared Function CheckAndGet(FixedRevUrl As String, KeyToCheck As String) As String

        'FixedRevUrl: Is the file where the information about the updates is.
        'CompletePath: The path will the file that will be splitted into folder and file
        'KeyToCheck: Is the key that will be checked in the file
        'Executable: Is the extracted file that will be executed, if Extract is set to false, this file will be the downloaded one

        Dim KeyFile As String = Path.GetTempPath & "\Dynawars-Key.txt"

        If File.Exists(KeyFile) Then
            File.Delete(KeyFile)
        End If

        My.Computer.Network.DownloadFile(
    FixedRevUrl,
    KeyFile)

        If IO.File.Exists(KeyFile) Then

            Dim Keys As String() = IO.File.ReadAllLines(KeyFile)

            If Not (Keys.Length = 0 Or Keys(0).Equals(String.Empty) Or Keys(1).Equals(String.Empty) Or WebUtils.CheckAddress(Keys(1))) Then

                If Not Keys(0) = KeyToCheck Then
                    Return Keys(1)
                End If

            End If

        End If

        Return Nothing

    End Function

    Public Shared Sub CheckAndDownload(FixedRevUrl As String, CompletePath As String, KeyToCheck As String, Optional CompressedExecutableName As String = "", Optional ExtractionPath As String = "")

        'FixedRevUrl: Is the file where the information about the updates is.
        'CompletePath: The path will the file that will be splitted into folder and file
        'KeyToCheck: Is the key that will be checked in the file
        'Executable: Is the extracted file that will be executed, if Extract is set to false, this file will be the downloaded one

        Dim RelativeFolder As String = CompletePath.Replace(CompletePath.Substring(CompletePath.LastIndexOf("\")), "")
        Dim RelativeFile As String = CompletePath.Replace(RelativeFolder, "")

        IOUtils.CreateDirectoryFromRoot(RelativeFolder)

        Dim KeyFile As String = RelativeFolder & "Key.txt"
        Dim DownloadedFile As String = ""
        Dim ExecutableFile As String = ""

        Dim Extract As Boolean = False
        Dim Execute As Boolean = False

        Dim FileExt As String = RelativeFile.Substring(RelativeFile.LastIndexOf(".") + 1)

        If FileExt.Equals("exe") Then
            Execute = True
        ElseIf FileExt.Equals("zip") Then
            Extract = True
        Else
            Throw New Exception("Este snippet solo acepta extensiones zip y exe")
            Exit Sub
        End If

        If Extract Then
            DownloadedFile = CompletePath
            If String.IsNullOrEmpty(CompressedExecutableName) Then
                ExecutableFile = RelativeFolder & "Setup.exe"
            Else
                ExecutableFile = RelativeFolder & CompressedExecutableName & ".exe"
            End If
        Else
            DownloadedFile = CompletePath
            ExecutableFile = CompletePath
        End If

        If File.Exists(KeyFile) Then
            File.Delete(KeyFile)
        End If

        My.Computer.Network.DownloadFile(
    FixedRevUrl,
    KeyFile)

        If IO.File.Exists(KeyFile) Then

            Dim Keys As String() = IO.File.ReadAllLines(KeyFile)

            If Not (Keys.Length = 0 Or Keys(0).Equals(String.Empty) Or Keys(1).Equals(String.Empty) Or WebUtils.CheckAddress(Keys(1))) Then

                If Extract And Not WebUtils.GetContentType(Keys(1)) = "application/zip" Then
                    MsgBox("El archivo descargado de la actualización tiene una extensión erronea.")
                    Exit Sub
                End If

                If Not Keys(0) = KeyToCheck Then
                    If MsgBox("¡Atención! Su aplicación está desactualizada." & vbCrLf & "Pulse ""Sí"" para continuar con la instalación. O ""No"" para descartar cambios.", MsgBoxStyle.YesNo, "¡Atención! Su app está desactualizada...") = MsgBoxResult.Yes Then

                        If File.Exists(DownloadedFile) Then
                            File.Delete(DownloadedFile)
                        End If

                        My.Computer.Network.DownloadFile(
                Keys(1),
                DownloadedFile)

                        If Extract Then

                            Dim ExtractionFolder As String = ""

                            If String.IsNullOrEmpty(ExtractionPath) Then
                                ExtractionFolder = RelativeFolder
                            Else
                                ExtractionFolder = ExtractionPath
                            End If

                            IOUtils.CreateDirectoryFromRoot(ExtractionFolder)

                            ZipManager.Extract(DownloadedFile, ExtractionFolder)

                        End If

                        If Execute Then
                            If Not File.Exists(ExecutableFile) Then
                                Throw New Exception("El archivo que se intento ejecutar no existe")
                                Exit Sub
                            End If
                            Dim psi As New ProcessStartInfo()
                            psi.UseShellExecute = True
                            psi.FileName = ExecutableFile
                            Process.Start(psi)
                            End
                        End If

                    End If

                Else

                    MsgBox("Hubo un error al extraer la información desde el archivo clave.", MsgBoxStyle.Critical, "Error")
                    Exit Sub

                End If

            End If

        Else

            MsgBox("Hubo un problema al descargar el archivo clave.", MsgBoxStyle.Critical, "Error")
            Exit Sub

        End If
    End Sub

End Class

Public Class WebUtils
    Public Shared Function CheckAddress(URL As String) As Boolean
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Dim response As WebResponse = request.GetResponse()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Shared Function GetContentType(URL As String) As String

        Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)

        ' Set some reasonable limits on resources used by this request
        request.MaximumAutomaticRedirections = 4
        request.MaximumResponseHeadersLength = 4

        ' Set credentials to use for this request.
        request.Credentials = CredentialCache.DefaultCredentials

        Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)

        Return response.ContentType

    End Function

End Class

Public Class IOUtils

    Private Shared ReadOnly Property OptionalPath
        Get
            Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\.dynawars\Game\"
        End Get
    End Property

    Public Shared Function NewGamePath(Desc As String, Optional OptPath As String = Nothing) As String

        If String.IsNullOrEmpty(OptPath) Then
            OptPath = OptionalPath
        End If

        Dim FolderPicker As New Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
        FolderPicker.ShowNewFolderButton = True

        FolderPicker.Description = Desc

        If FolderPicker.ShowDialog.ToString() = "True" Then
            'Console.WriteLine("Carpeta: " & FolderPicker.SelectedPath & ", resultado: " & FolderPicker.ShowDialog.ToString())
            Return FolderPicker.SelectedPath
        Else
            MsgBox("La ruta por defecto en el que el juego se instalará es """ & OptPath & """", MsgBoxStyle.Information, "Información")
            'Console.WriteLine("Carpeta: " & FolderPicker.SelectedPath & ", resultado: " & FolderPicker.ShowDialog.ToString())
            Return OptPath
        End If

    End Function

    Public Shared Sub CreateDirectoryFromRoot(path As String)

        Dim array As String() = path.Split("\")
        Dim subfolders As Integer = array.Length - 1
        Dim indexbef As String = ""

        For i As Integer = 0 To subfolders
            indexbef += array(i) & "\"
            If Not Directory.Exists(indexbef) Then
                Directory.CreateDirectory(indexbef)
            End If
        Next

    End Sub

    Public Shared Function ValidatePath(path As String) As String
        If path.Contains("/") Then
            If path.Substring(path.LastIndexOf("/") - 1, 1) IsNot "/" Then
                path += "/"
            End If
        ElseIf path.Contains("\") Then
            If path.Substring(path.LastIndexOf("\") - 1, 1) IsNot "\" Then
                path += "\"
            End If
        End If
        Return path
    End Function

End Class

Public Class ProfileItem
    Public Property ProfileName() As String
        Get
            Return m_ProfileName
        End Get
        Set(value As String)
            m_ProfileName = value
        End Set
    End Property
    Private m_ProfileName As String
    Public Property VersionName() As String
        Get
            Return m_VersionName
        End Get
        Set(value As String)
            m_VersionName = value
        End Set
    End Property
    Private m_VersionName As String
    Public Property Username() As String
        Get
            Return m_Username
        End Get
        Set(value As String)
            m_Username = value
        End Set
    End Property
    Private m_Username As String
End Class

Public Class KeyBoardItem
    Public Property Key() As String
        Get
            Return m_Key
        End Get
        Set(value As String)
            m_Key = value
        End Set
    End Property
    Private m_Key As String
    Public Property Action() As String
        Get
            Return m_Action
        End Get
        Set(value As String)
            m_Action = value
        End Set
    End Property
    Private m_Action As String
End Class

Public Class Profile

    Public Shared LoadedProfile As String = ""

    Private Shared _isLoggedIn As Boolean = False
    Public Shared Property IsLoggedIn() As Boolean
        Get
            Return _isLoggedIn
        End Get
        Set(value As Boolean)
            _isLoggedIn = value
        End Set
    End Property

    Private Shared _username As String
    Public Shared Property UserName() As String
        Get
            Return _username
        End Get
        Set(value As String)
            _username = value
        End Set
    End Property

End Class

Public Class Paths
    Public Shared GamePath As String = ""
    Public Shared AppPath As String = ""
End Class

Public Class Win32
    ' Methods
    <DllImport("user32.dll")> _
    Friend Shared Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As POINT) As Boolean
    End Function

    <DllImport("user32.dll")> _
    Friend Shared Function MoveWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function


    ' Nested Types
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure POINT
        Public X As Integer
        Public Y As Integer
        Public Sub New(ByVal x As Integer, ByVal y As Integer)
            Me.X = x
            Me.Y = y
        End Sub

        Public Sub New(ByVal pt As Windows.Point)
            Me.X = Convert.ToInt32(pt.X)
            Me.Y = Convert.ToInt32(pt.Y)
        End Sub
    End Structure
End Class

Public Class Launcher

    Private Shared mw As MainWindow = Application.Current.Windows.OfType(Of MainWindow)().FirstOrDefault()

    'Private Shared ReadOnly Property mw
    '    Get
    '        Return Application.Current.Windows.OfType(Of MainWindow)().FirstOrDefault()
    '    End Get
    'End Property

    Public Shared controls As Collection(Of ControlsItems)

    Public Shared Sub CreateFolders()

        IOUtils.CreateDirectoryFromRoot(MainWindow.AppWorkPath & "Launcher\Settings")

        If MainWindow.IsGameInstalled Then
            IOUtils.CreateDirectoryFromRoot(MainWindow.GamePath & "Game_Data\Profiles")
        End If

    End Sub

    Public Shared Sub LoadAllProfiles()

        If String.IsNullOrEmpty(Profile.UserName) Then
            MsgBox("Debes loguearte/registrarte antes de poder crear un perfil.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        If mw.dtProfiles.Columns Is Nothing Then

            Dim c1 As New DataGridTextColumn()
            c1.Header = "Nombre"
            c1.Binding = New Binding("ProfileName")
            c1.IsReadOnly = True
            c1.Width = 241
            mw.dtProfiles.Columns.Add(c1)

            Dim c2 As New DataGridTextColumn()
            c2.Header = "Versión"
            c2.Binding = New Binding("VersionName")
            c2.IsReadOnly = True
            c2.Width = 241
            mw.dtProfiles.Columns.Add(c2)

            Dim c3 As New DataGridTextColumn()
            c3.Header = "Usuario"
            c3.Binding = New Binding("Username")
            c3.IsReadOnly = True
            c3.Width = 241
            mw.dtProfiles.Columns.Add(c3)

        End If

        For Each folder As String In Directory.GetDirectories(MainWindow.GamePath & "Game_Data\Profiles")

            INIFileManager.FilePath = IO.Path.Combine(folder & "\", "profile.cfg")
            mw.cmbProfiles.Items.Add(INIFileManager.Key.Get("ProfileName"))

            mw.dtProfiles.Items.Add(New ProfileItem() With { _
                .ProfileName = INIFileManager.Key.Get("ProfileName"), _
                .VersionName = INIFileManager.Key.Get("VersionName"), _
                .Username = INIFileManager.Key.Get("UserName") _
            })

        Next

    End Sub

    Public Shared Sub LoadProfile(Name As String)

        If String.IsNullOrEmpty(Profile.UserName) Then
            MsgBox("Debes loguearte/registrarte antes de poder cargar cualquier perfil.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        mw.Profile_Name.Text = Name

        INIFileManager.FilePath = Path.Combine(MainWindow.GamePath & "Game_Data\Profiles\" & Name.ToLower, "\profile.cfg")
        mw.Sensibility.Value = INIFileManager.Key.Get("Sensibility")
        mw.FOV.Value = INIFileManager.Key.Get("FOV")
        mw.Brightness.Value = Convert.ToInt32(Convert.ToDouble(INIFileManager.Key.Get("Brightness").ToString.Replace(".", ",")) * 100)
        mw.RenderDistance.Value = INIFileManager.Key.Get("RenderDistance")

        If INIFileManager.Key.Get("MaxFPS") = -1 Then
            mw.MaxFPS.Value = 200
        ElseIf INIFileManager.Key.Get("VSync") = 1 Then
            mw.MaxFPS.Value = 5
        Else
            mw.MaxFPS.Value = INIFileManager.Key.Get("MaxFPS")
        End If

        Dim fov_res As Double = ((mw.FOV.Value - 60) \ (120 - 60)) * 200

        If mw.FOV.Value = 60 Then
            mw.FOV_all = "Normal"
        ElseIf mw.FOV.Value = 120 Then
            mw.FOV_all = "Quake Pro"
        Else
            mw.FOV_all = fov_res.ToString("F0")
        End If

        Dim brightness_res As Double = ((mw.Brightness.Value - 1) \ (30 - 1)) * 100

        If mw.Brightness.Value = 30 Then
            mw.Brightness_all = "Claro"
        ElseIf mw.Brightness.Value = 1 Then
            mw.Brightness_all = "Oscuro"
        Else
            mw.Brightness_all = brightness_res.ToString("F0")
        End If

        If mw.MaxFPS.Value = 200 Then
            mw.Maxfps_all = "A tope"
        ElseIf mw.MaxFPS.Value = 5 Then
            mw.Maxfps_all = "VSync"
        Else
            mw.Maxfps_all = mw.MaxFPS.Value.ToString("F0")
        End If

        Dim sens_res As Double = ((mw.Sensibility.Value - 0.1) \ (30 - 0.1)) * 200

        If mw.Sensibility.Value = 30 Then
            mw.Sens_All = "HIPERVELOCIDAD!"
        ElseIf mw.Sensibility.Value = 0.1 Then
            mw.Sens_All = "*bostezo*"
        Else
            mw.Sens_All = sens_res.ToString("F0")
        End If

        mw.lblFOV.Content = "FOV (Field Of View) [" & mw.FOV_all & "]:"
        mw.lblSens.Content = "Sensibilidad [" & mw.Sens_All & "]:"
        mw.lblMaxFPS.Content = "FPS Máximos [" & mw.Maxfps_all & "]:"
        mw.lblRenderDistance.Content = "FOV (Field Of View) [" & (mw.RenderDistance.Value \ 1000).ToString("F1") & " Km]:"
        mw.lblBrightness.Content = "Brillo [" & mw.Brightness_all & "]:"

        Profile.LoadedProfile = Name.ToLower

        LoadAllControls()

    End Sub

    Public Shared Sub SaveProfile()

        If String.IsNullOrEmpty(Profile.UserName) Then
            MsgBox("Debes loguearte/registrarte antes de poder guardar un perfil.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        Dim alreadyCreated As Boolean = True

        If String.IsNullOrEmpty(mw.Profile_Name.Text) Then
            MsgBox("Debes poner un nombre al perfil para poder guardarlo.", MsgBoxStyle.Exclamation, "Advertencia")
            Exit Sub
        End If

        'Check if the profile was not created and create one

        If mw.dtControls.Items.Count = 0 Then 'This mean that no control has been loaded so create a new profile
            CreateProfile()
        End If

        If Not Directory.Exists(MainWindow.GamePath & "Game_Data\Profiles\" & mw.Profile_Name.Text.ToLower) Then
            Directory.CreateDirectory(MainWindow.GamePath & "Game_Data\Profiles\" & mw.Profile_Name.Text.ToLower)
        End If

        INIFileManager.FilePath = MainWindow.GamePath & "Game_Data\Profiles\" & mw.Profile_Name.Text.ToLower & "\profile.cfg"

        If Not File.Exists(MainWindow.GamePath & "Game_Data\Profiles\" & mw.Profile_Name.Text.ToLower & "\profile.cfg") Then
            INIFileManager.File.Create()
            alreadyCreated = False
        End If

        Dim maxfps_all As Int32 = mw.MaxFPS.Value
        Dim vsync_all As Boolean = False

        If mw.MaxFPS.Value = 200 Then
            maxfps_all = -1
        ElseIf mw.MaxFPS.Value = 5 Then
            vsync_all = True
            maxfps_all = -1
        End If

        INIFileManager.Key.Set("ProfileName", mw.Profile_Name.Text)
        INIFileManager.Key.Set("VersionName", mw.Profile_Name.Text.ToLower & "-0.0.1 WIP")
        INIFileManager.Key.Set("UserName", Profile.UserName)
        INIFileManager.Key.Set("FOV", mw.FOV.Value)
        INIFileManager.Key.Set("Sensibility", mw.Sensibility.Value)
        INIFileManager.Key.Set("Brightness", (mw.Brightness.Value \ 100).ToString.Replace(",", "."))
        INIFileManager.Key.Set("RenderDistance", mw.RenderDistance.Value)
        INIFileManager.Key.Set("MaxFPS", maxfps_all)
        INIFileManager.Key.Set("InvertMouse", Convert.ToInt32(mw.InvertMouse.IsChecked).ToString())
        INIFileManager.Key.Set("VSync", Convert.ToInt32(vsync_all).ToString())
        INIFileManager.Key.Set("QualityLvl", Convert.ToInt32(mw.QualityLvl.Value).ToString())

        'Save fullScreen
        INIFileManager.Key.Set("fullScreen", Convert.ToInt32(mw.chkFullScreen.IsChecked))

        'Save controls

        Dim controlsDict As Dictionary(Of String, String) = New Dictionary(Of String, String)

        For Each item As ControlsItems In controls
            controlsDict.Add(item.Action, item.Key)
        Next

        INIUtils.Create_Group(INIFileManager.FilePath, "Keys", controlsDict)

        'Set & save resolution

        Dim winResH As Double = System.Windows.SystemParameters.PrimaryScreenHeight
        Dim winResW As Double = System.Windows.SystemParameters.PrimaryScreenWidth

        Dim preH As String = INIFileManager.Key.Get("gameH")
        Dim preW As String = INIFileManager.Key.Get("gameW")

        Dim HeightRes As Double
        Dim WidthRes As Double

        If mw.ComboBoxH.IsVisible Then
            HeightRes = Double.Parse(DirectCast(mw.ComboBoxH.SelectedItem, ComboBoxItem).Content.ToString())
        ElseIf mw.txtBoxH.IsVisible Then
            HeightRes = Double.Parse(mw.txtBoxH.Text)
        End If

        If mw.ComboBoxH.IsVisible Then
            WidthRes = Double.Parse(DirectCast(mw.ComboBoxW.SelectedItem, ComboBoxItem).Content.ToString())
        ElseIf mw.txtBoxH.IsVisible Then
            WidthRes = Double.Parse(mw.txtBoxW.Text)
        End If

        If (mw.ComboBoxH.SelectedIndex = -1 And (mw.txtBoxH.Text = "Altura" Or Not mw.txtBoxH.IsVisible) And String.IsNullOrEmpty(preH)) Or (mw.ComboBoxW.SelectedIndex = -1 And (mw.txtBoxW.Text = "Anchura" Or Not mw.txtBoxW.IsVisible) And String.IsNullOrEmpty(preW)) Then

            If mw.ComboBoxH.SelectedIndex = -1 And (mw.txtBoxH.Text = "Altura" Or Not mw.txtBoxH.IsVisible Or String.IsNullOrEmpty(mw.txtBoxH.Text)) And String.IsNullOrEmpty(preH) And mw.ComboBoxW.SelectedIndex = -1 And (mw.txtBoxW.Text = "Anchura" Or Not mw.txtBoxW.IsVisible Or String.IsNullOrEmpty(mw.txtBoxW.Text)) And String.IsNullOrEmpty(preW) Then

                If MsgBox("No hay ninguna resolución establecida con anterioridad. ¿Deseas que la resolución sea de tu pantalla, es decir, " & winResW.ToString() & "x" & winResH.ToString() & "?", MsgBoxStyle.YesNo, "Información") = MsgBoxResult.Yes Then
                    HeightRes = winResH
                    WidthRes = winResW
                    'ANOTACIÓN DE LÓGICA: Else + exit if ni hace falta puesto que elseif no va ser leido...
                Else
                    Dim strRes As String = InputBox("Introduzca una resolución, ej: ""1280x720""", "Introduzca la resolución", winResW.ToString() & "x" & winResH.ToString())
                    HeightRes = Convert.ToDouble(strRes.Split("x").Last())
                    WidthRes = Convert.ToDouble(strRes.Split("x").First())
                End If

            ElseIf mw.ComboBoxH.SelectedIndex = -1 And (mw.txtBoxH.Text = "Altura" Or Not mw.txtBoxH.IsVisible Or String.IsNullOrEmpty(mw.txtBoxH.Text)) And String.IsNullOrEmpty(preH) Then

                If MsgBox("No hay ninguna altura establecida con anterioridad. ¿Deseas que la altura sea de tu pantalla, es decir, " & winResH.ToString() & "?", MsgBoxStyle.YesNo, "Información") = MsgBoxResult.Yes Then
                    HeightRes = winResH
                Else
                    HeightRes = Convert.ToDouble(InputBox("Introduzca una altura como resolución", "Introduzca la resolución", winResH.ToString()))
                End If

            ElseIf mw.ComboBoxW.SelectedIndex = -1 And (mw.txtBoxW.Text = "Anchura" Or Not mw.txtBoxW.IsVisible Or String.IsNullOrEmpty(mw.txtBoxW.Text)) And String.IsNullOrEmpty(preW) Then

                If MsgBox("No hay ninguna resolución establecida con anterioridad. ¿Deseas que la resolución sea de tu pantalla, es decir, " & winResW.ToString() & "x" & winResH.ToString() & "?", MsgBoxStyle.YesNo, "Información") = MsgBoxResult.Yes Then
                    WidthRes = winResW
                Else
                    WidthRes = Convert.ToDouble(InputBox("Introduzca una anchura como resolución", "Introduzca la resolución", winResW.ToString()))
                End If

            End If

        ElseIf DirectCast(mw.ComboBoxH.SelectedItem, ComboBoxItem).Content.ToString().Equals(preH) Or DirectCast(mw.ComboBoxW.SelectedItem, ComboBoxItem).Content.ToString().Equals(preW) Then

            If DirectCast(mw.ComboBoxH.SelectedItem, ComboBoxItem).Content.ToString().Equals(preH) And DirectCast(mw.ComboBoxW.SelectedItem, ComboBoxItem).Content.ToString().Equals(preW) Then
                GoTo SkipRes
            ElseIf DirectCast(mw.ComboBoxH.SelectedItem, ComboBoxItem).Content.ToString().Equals(preH) Then
                GoTo SkipRes
            ElseIf DirectCast(mw.ComboBoxW.SelectedItem, ComboBoxItem).Content.ToString().Equals(preW) Then
                GoTo SkipRes
            End If

        ElseIf mw.txtBoxH.Text.Equals(preH) Or mw.txtBoxW.Text.Equals(preW) Then

            If mw.txtBoxH.Text.Equals(preH) And mw.txtBoxW.Text.Equals(preW) Then
                GoTo SkipRes
            ElseIf mw.txtBoxH.Text.Equals(preH) Then
                GoTo SkipRes
            ElseIf mw.txtBoxW.Text.Equals(preW) Then
                GoTo SkipRes
            End If

        End If

        Dim dangerous As Boolean = False

        If winResW < WidthRes Or winResH < HeightRes Then

            If winResW < WidthRes And winResH < HeightRes Then

                If MsgBox("Ambas resoluciones son mayores que la que tu posees, ¿deseas continuar?", MsgBoxStyle.YesNo, "Advertencia") = MsgBoxResult.Yes Then
                    dangerous = True
                Else
                    GoTo SkipRes
                End If

            Else

                If winResW < WidthRes Then

                    If MsgBox("La anchura de la resolución que tu seleccionaste es mayor de la que tu posees, ¿deseas continuar?", MsgBoxStyle.YesNo, "Advertencia") = MsgBoxResult.Yes Then
                        dangerous = True
                    Else
                        GoTo SkipRes
                    End If

                ElseIf winResH < HeightRes Then

                    If MsgBox("La altura de la resolución que tu seleccionaste es mayor de la que tu posees, ¿deseas continuar?", MsgBoxStyle.YesNo, "Advertencia") = MsgBoxResult.Yes Then
                        dangerous = True
                    Else
                        GoTo SkipRes
                    End If

                End If

            End If

        Else

            dangerous = True

        End If

        If dangerous Then
            INIFileManager.Key.Set("app-Width", WidthRes.ToString())
            INIFileManager.Key.Set("app-Height", HeightRes.ToString())
        End If

SkipRes:

        LoadAllProfiles()

        UnloadAll() 'Unload all...

        If Not alreadyCreated Then
            MsgBox("Perfil creado. En """ & INIFileManager.FilePath & """", MsgBoxStyle.Information, "Información")
        Else
            MsgBox("Perfil editado.", MsgBoxStyle.Information, "Información")
        End If

    End Sub

    Public Shared Sub CreateProfile()

        If String.IsNullOrEmpty(Profile.UserName) Then
            MsgBox("Debes loguearte/registrarte antes de poder crear un perfil.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        mw.FOV.Value = 60
        mw.Sensibility.Value = 15
        mw.Brightness.Value = 0.3
        mw.RenderDistance.Value = 1500
        mw.MaxFPS.Value = 5
        mw.InvertMouse.IsChecked = False

        CreateAllControls()

    End Sub

    Public Shared Sub UnloaAllControls()

        mw.FOV.Value = 0
        mw.Sensibility.Value = 0
        mw.Brightness.Value = 0
        mw.RenderDistance.Value = 0
        mw.MaxFPS.Value = 0
        mw.InvertMouse.IsChecked = False

    End Sub

    Public Shared Sub UnloadAll()

        UnloadAllControls()
        UnloadAllControls()

        'mw.dtControls.Items.Add(New ControlsItems() With { _
        '    .Action = "", _
        '    .Key = "" _
        '})

    End Sub

    Public Shared Sub LoadAllLauncherSettings()

        INIFileManager.FilePath = Path.Combine(MainWindow.AppWorkPath & "Launcher\Settings\", "launcherConfig.cfg")

        mw.ComboBoxW.SelectedItem = INIFileManager.Key.Get("app-Width")
        mw.ComboBoxH.SelectedItem = INIFileManager.Key.Get("app-Height")
        mw.chkFullScreen.IsChecked = Convert.ToBoolean(INIFileManager.Key.Get("fullScreen"))

    End Sub

    Public Shared Sub SaveLauncherSettings()

        'Lenguaje y aparencia

    End Sub

    Public Shared Sub Log_in()

        Dim sourcecode As String = New System.Net.WebClient().DownloadString("http://dynawars.x10host.com/gameact.php?go=login&action=check&user=" & mw.user.Text & "&pass=" & mw.pass.Password)

        'Login here
        Select Case sourcecode

            Case "1"
                Profile.IsLoggedIn = True
                Profile.UserName = mw.user.Text
                mw.username.Text = Profile.UserName
                mw.loginPanel.Visibility = Windows.Visibility.Hidden
                mw.loggedPanel.Visibility = Windows.Visibility.Visible

                Launcher.LoadAllProfiles()

            Case Else

                If sourcecode.Substring(0, 8).Equals("Error #0") Then
                    If MsgBox("El usuario no fue encontrado en nuestras bases de datos. ¿Deseas registrar una nueva cuenta?", MsgBoxStyle.YesNo, "Error") = MsgBoxResult.Yes Then
                        mw.wbWindow.WebBrowser.Navigate("http://dynawars.x10host.com/web/user/register")
                    End If
                ElseIf sourcecode.Substring(0, 8).Equals("Error #1") Then
                    MsgBox("Parece que hubo un error interno, intentalo más tarde.")
                ElseIf sourcecode.Substring(0, 8).Equals("Error #2") Then
                    If MsgBox("Parece que te equivocaste de contraseña. ¿La olvidaste?", MsgBoxStyle.YesNo, "Error") = MsgBoxResult.Yes Then
                        mw.wbWindow.WebBrowser.Navigate("http://dynawars.x10host.com/web/user/password")
                    End If
                End If

        End Select

    End Sub

    Public Shared Sub Log_out()

        Profile.IsLoggedIn = False
        Profile.UserName = ""

        mw.loginPanel.Visibility = Windows.Visibility.Visible
        mw.loggedPanel.Visibility = Windows.Visibility.Hidden

        mw.user.Text = ""
        mw.pass.Password = ""

    End Sub

    Public Shared Sub CreateAllControls()

        If mw.dtControls.Columns Is Nothing Then

            Dim c4 As New DataGridTextColumn()
            c4.Header = "Acción"
            c4.Binding = New Binding("Action")
            c4.IsReadOnly = True
            c4.Width = 241
            mw.dtControls.Columns.Add(c4)

            Dim c5 As New DataGridTextColumn()
            c5.Header = "Tecla/Botón"
            c5.Binding = New Binding("Key")
            c5.IsReadOnly = True
            c5.Width = 241
            mw.dtControls.Columns.Add(c5)

        End If

        controls = New Collection(Of ControlsItems)

        'Dim tmpInt As Integer = 0

        For Each item As KeyValuePair(Of String, String) In GameApi.Controls.DefaultInputMap

            mw.dtControls.Items.Add(New ControlsItems() With { _
                .Action = item.Key, _
                .Key = item.Value _
            })

            'Dim rowNum As Integer = Integer.Parse(Math.Floor(tmpInt / mw.dtControls.Columns.Count))
            'Dim columnNum As Integer = Integer.Parse(Math.Floor(((tmpInt / mw.dtControls.Columns.Count) - rowNum) * mw.dtControls.Columns.Count))

            'DataGridHelper.GetCell(mw.dtControls, rowNum, columnNum)

            'tmpInt += 1

            controls.Add(New ControlsItems() With { _
                .Action = item.Key, _
                .Key = item.Value _
            })

        Next

    End Sub

    Public Shared Sub LoadAllControls()

        If String.IsNullOrEmpty(Profile.LoadedProfile) Then
            Throw New Exception("No se ha cargado ningún perfil")
            Exit Sub
        End If

        If mw.dtControls.Columns(0).Header.ToString() IsNot "Acción" Then
            Dim c4 As New DataGridTextColumn()
            c4.Header = "Acción"
            c4.Binding = New Binding("Action")
            c4.IsReadOnly = True
            c4.Width = 241
            mw.dtControls.Columns.Add(c4)
        End If

        If mw.dtControls.Columns(1).Header.ToString() IsNot "Tecla/Botón" Then
            Dim c5 As New DataGridTextColumn()
            c5.Header = "Tecla/Botón"
            c5.Binding = New Binding("Key")
            c5.IsReadOnly = True
            c5.Width = 241
            mw.dtControls.Columns.Add(c5)
        End If

        controls = New Collection(Of ControlsItems)

        For Each item As KeyValuePair(Of String, String) In GameApi.Controls.InputMap

            mw.dtControls.Items.Add(New ControlsItems() With { _
                .Action = item.Key, _
                .Key = item.Value _
            })

            controls.Add(New ControlsItems() With { _
                .Action = item.Key, _
                .Key = item.Value _
            })

        Next

    End Sub

    Public Shared Sub UnloadAllControls()

        GameApi.Controls.InputMap = Nothing

        DataGridUtils.ClearAll(mw.dtControls)

    End Sub

    Public Shared Sub SaveAllControls()

        If String.IsNullOrEmpty(Profile.LoadedProfile) Then
            Throw New Exception("No se ha cargado ningún perfil")
            Exit Sub
        End If

        Dim KeysFromGrid As Dictionary(Of String, String) = New Dictionary(Of String, String)

        For i As Integer = 0 To Launcher.controls.Count
            KeysFromGrid.Add(Launcher.controls.Item(i).Action, Launcher.controls.Item(i).Key)
        Next

        INIUtils.Edit_Group(MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg", "Keys", KeysFromGrid)

    End Sub

    Public Shared Sub ShowTuto()

        If MainWindow.TutoShown Then
            Exit Sub
        End If

        MainWindow.TutoShown = True

        If File.Exists(MainWindow.AppWorkPath & "Launcher\Settings\" & "launcherConfig.cfg") Then
            INIFileManager.FilePath = MainWindow.AppWorkPath & "Launcher\Settings\" & "launcherConfig.cfg"
        Else
            IO.File.Create(MainWindow.AppWorkPath & "Launcher\Settings\" & "launcherConfig.cfg")
            'INIFileManager.File.Create()
            INIFileManager.FilePath = MainWindow.AppWorkPath & "Launcher\Settings\" & "launcherConfig.cfg"
            INIFileManager.Key.Set("TutoIni", 0)
        End If

        Dim tutoPrincipal As Boolean = Convert.ToBoolean(Convert.ToInt32(INIFileManager.Key.Get("TutoIni")))

        If Not tutoPrincipal Then

            If MsgBox("¿Deseas ver el archivo que contiene el tutorial inicial? (Pulse No para no mostrar más)", MsgBoxStyle.YesNoCancel, "?") = MsgBoxResult.Yes Then
                MsgBox("OPCIÓN NO DISPONIBLE", MsgBoxStyle.Critical, "Error")
                Exit Sub
            ElseIf MessageBoxResult.No Then
                INIFileManager.Key.Set("TutoIni", 1)
                Exit Sub
            ElseIf MessageBoxResult.Cancel Then
                Exit Sub
            End If
        End If

    End Sub

End Class

Public Class SpecialTextFormat

    Public Shared Function FormatSeconds(iSecond As Double) As String

        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(iSecond)

        If iSecond < 60 Then

            Return iSpan.Seconds.ToString.PadLeft(2, "0"c) & " s"

        ElseIf iSecond < 3600 Then

            Return iSpan.Minutes.ToString.PadLeft(2, "0"c) & " m " & _
                iSpan.Seconds.ToString.PadLeft(2, "0"c) & " s"

        Else

            Return iSpan.Hours.ToString.PadLeft(2, "0"c) & " h " & _
                iSpan.Minutes.ToString.PadLeft(2, "0"c) & " m " & _
                iSpan.Seconds.ToString.PadLeft(2, "0"c) & " s"

        End If

    End Function

    Public Shared Function FormatBytes(MIB As Integer) As String

        Dim strMIB As String = MIB.ToString("0.00") & " B"

        If MIB > 1000 Then
            Return (MIB \ 1000).ToString("0.00") & " KB"
        ElseIf MIB > 1000000 Then
            Return (MIB \ 1000000).ToString("0.00") & " B"
        ElseIf MIB > 1000000000 Then
            Return (MIB \ 1000000000).ToString("0.00") & " GB"
        Else
            Return strMIB
        End If

    End Function

End Class

Public Class DataGridUtils

    Private Shared Sub InternalDelete(dataGrid1 As DataGrid)

        For i As Integer = 0 To dataGrid1.Items.Count / dataGrid1.Columns.Count - 1

            Dim IsColumnEmpty(dataGrid1.Columns.Count - 1) As Boolean

            For c As Integer = 0 To dataGrid1.Columns.Count - 1
                IsColumnEmpty(c) = Not String.IsNullOrEmpty(DataGridHelper.GetCell(dataGrid1, i, c).Content.ToString())
            Next

            If IsColumnEmpty.All(Function(x) x = True) Then
                dataGrid1.Items.RemoveAt(i)
            End If

        Next

    End Sub

    Public Shared Sub DeleteEmptyRows(dataGrid1 As DataGrid)

        'Wait a prudent time until everything is rendered, for example 10ms, for avoid NullRefenceException
        Dim timer As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer() With {
                .Interval = TimeSpan.FromSeconds(0.01D)
            }

        timer.Start()

        AddHandler timer.Tick, Sub()
                                   timer.Stop()
                                   InternalDelete(dataGrid1)
                               End Sub



    End Sub

    Public Shared Sub ClearAll(dataGrid1 As DataGrid)

        For i As Integer = 0 To dataGrid1.Items.Count / dataGrid1.Columns.Count - 1

            dataGrid1.Items.RemoveAt(i)

        Next

    End Sub

End Class

Public Class DataGridHelper
    Public Shared Function GetCell(dg As DataGrid, row As Integer, column As Integer) As DataGridCell

        Dim rowContainer As DataGridRow = GetRow(dg, row)

        If rowContainer IsNot Nothing Then
            Dim presenter As Primitives.DataGridCellsPresenter = GetVisualChild(Of Primitives.DataGridCellsPresenter)(rowContainer)

            ' try to get the cell but it may possibly be virtualized
            Dim cell As DataGridCell = DirectCast(presenter.ItemContainerGenerator.ContainerFromIndex(column), DataGridCell)
            If cell Is Nothing Then
                ' now try to bring into view and retreive the cell
                dg.ScrollIntoView(rowContainer, dg.Columns(column))
                cell = DirectCast(presenter.ItemContainerGenerator.ContainerFromIndex(column), DataGridCell)
            End If
            Return cell
        End If

        Return Nothing

    End Function

    Public Shared Function GetRow(dg As DataGrid, index As Integer) As DataGridRow

        Dim row As DataGridRow = DirectCast(dg.ItemContainerGenerator.ContainerFromIndex(index), DataGridRow)

        If row Is Nothing Then
            ' may be virtualized, bring into view and try again
            dg.ScrollIntoView(dg.Items(index))
            row = DirectCast(dg.ItemContainerGenerator.ContainerFromIndex(index), DataGridRow)
        End If

        Return row

    End Function

    Public Shared Function TryToFindGridCell(grid As DataGrid, cellInfo As DataGridCellInfo) As DataGridCell
        Dim result As DataGridCell = Nothing
        Dim row As DataGridRow = DirectCast(grid.ItemContainerGenerator.ContainerFromItem(cellInfo.Item), DataGridRow)
        If row IsNot Nothing Then
            Dim columnIndex As Integer = grid.Columns.IndexOf(cellInfo.Column)
            If columnIndex > -1 Then
                Dim presenter As Primitives.DataGridCellsPresenter = GetVisualChild(Of Primitives.DataGridCellsPresenter)(row)
                result = TryCast(presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex), DataGridCell)
            End If
        End If
        Return result
    End Function

    Private Shared Function GetVisualChild(Of T As Visual)(parent As Visual) As T
        Dim child As T = Nothing
        Dim numVisuals As Integer = VisualTreeHelper.GetChildrenCount(parent)
        For i As Integer = 0 To numVisuals - 1
            Dim v As Visual = DirectCast(VisualTreeHelper.GetChild(parent, i), Visual)
            child = TryCast(v, T)
            If child Is Nothing Then
                child = GetVisualChild(Of T)(v)
            End If
            If child IsNot Nothing Then
                Exit For
            End If
        Next
        Return child
    End Function

End Class

Namespace GameApi

    Public Class Controls

        Private Shared m_inputMap As Dictionary(Of String, String)

        Private Shared defInputMap As Dictionary(Of String, String)

        Private Shared dftKMStrArray As String()

        Public Shared Property InputMap() As Dictionary(Of String, String)
            Get
                If m_inputMap Is Nothing Then
                    Dim tmp = LoadAllKeysFromConfig()
                    If tmp Is Nothing Then
                        Return LoadAllDefaultKeys()
                    End If
                    Return tmp
                End If
                Return m_inputMap
            End Get
            Set(value As Dictionary(Of String, String))
                m_inputMap = value
            End Set
        End Property

        Public Shared Property DefaultInputMap() As Dictionary(Of String, String)
            Get
                If defInputMap Is Nothing Then
                    Return LoadAllDefaultKeys()
                End If
                Return defInputMap
            End Get
            Set(value As Dictionary(Of String, String))
                defInputMap = value
            End Set
        End Property

        Public Shared ReadOnly Property DefaultKeysActions() As String()
            Get

                Dim tempdefstrarray As New List(Of String)()

                For Each kvpDefMap As KeyValuePair(Of String, String) In DefaultInputMap
                    tempdefstrarray.Add(kvpDefMap.Key)
                Next

                Return tempdefstrarray.ToArray()
            End Get
        End Property

        Public Shared Function LoadAllKeysFromConfig() As Dictionary(Of String, String)

            If Not IniUtils.Group_Exists(MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg", "Keys") Then
                'Dim nl As String = System.Environment.NewLine
                'Dim InitialLoadAllKeysFromConfig As String = "Exp=L" & nl & "Kill=K" & nl & "FPS=F10" & nl & "EnableGUI=F1" & nl & "GameMenu=Escape" & nl & "Inv=E" & nl & "Waypoint=P" & nl & "ToggleMap=M" & nl & "SwitchView=V" & nl & "TakeScreenShot=F2" & nl & "Laser=L" & nl & "FlashLight=F" & nl & "Reload=R" & nl & "Console=F12" & nl & "SendCommand=Return" & nl & "Unlock=U"
                INIUtils.Create_Group(MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg", "Keys", DefaultInputMap) 'InitialLoadAllKeysFromConfig
            End If

            Dim GetKeys As String = IniUtils.Read_Group("Keys", MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg")
            Dim tempInputMapRaw As Dictionary(Of String, String) = IniUtils.Extract_Keys("", GetKeys)
            Dim tempInputMap As New Dictionary(Of String, String)()
            Dim tmpImportedKeys As New List(Of String)()
            For Each kvpInputRaw As KeyValuePair(Of String, String) In tempInputMapRaw
                tempInputMap.Add(kvpInputRaw.Key, kvpInputRaw.Value)
                tmpImportedKeys.Add(kvpInputRaw.Key)
            Next

            If tmpImportedKeys.ToArray().Length <> DefaultKeysActions.Length Then
                'Check if some index is missing
                Dim missingIndexes As String() = tmpImportedKeys.ToArray().Except(DefaultKeysActions).ToArray()

                For Each missingIndex As String In missingIndexes
                    tempInputMap.Add(missingIndex, GetKeyFromDefLib(missingIndex))
                Next
            End If

            Return tempInputMap

        End Function

        Public Shared Function LoadAllDefaultKeys() As Dictionary(Of String, String)

            defInputMap = New Dictionary(Of String, String)()

            defInputMap.Add("Exp", "L")

            defInputMap.Add("Kill", "K")

            defInputMap.Add("FPS", "F10")

            defInputMap.Add("EnableGUI", "F1")

            defInputMap.Add("GameMenu", "Escape")

            defInputMap.Add("Inv", "E")

            defInputMap.Add("Waypoint", "P")

            defInputMap.Add("ToggleMap", "M")

            defInputMap.Add("SwitchView", "V")

            defInputMap.Add("TakeScreenShot", "F2")

            defInputMap.Add("Laser", "L")

            defInputMap.Add("FlashLight", "F")

            defInputMap.Add("Reload", "R")

            defInputMap.Add("Console", "F12")

            defInputMap.Add("SendCommand", "Return")

            defInputMap.Add("Unlock", "U")

            Return defInputMap

        End Function

        Public Shared Function GetKeyFromDefLib(actionName As String) As String
            Return DefaultInputMap(actionName)
        End Function

        Public Shared Sub RestoreAll()
            InputMap = DefaultInputMap
        End Sub

        Public Shared Sub RestoreKey(actionName As String)
            InputMap(actionName) = DefaultInputMap(actionName)
        End Sub

        Public Shared Function GetKey(action As String) As String

            If String.IsNullOrEmpty(Profile.LoadedProfile) Then
                Throw New Exception("No se ha cargado ningún perfil")
                Exit Function
            End If

            If INIUtils.Group_Exists(MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg", "Keys") Then
                Dim GetKeys As String = INIUtils.Read_Group("Keys", MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile & "\profile.cfg")
                Dim tempInputMapRaw As Dictionary(Of String, String) = INIUtils.Extract_Keys("", GetKeys)
                Return tempInputMapRaw(action)
            End If
            Return Nothing
        End Function

    End Class

End Namespace

Public Class ControlsItems
    Public Property Key() As String
        Get
            Return m_Key
        End Get
        Set(value As String)
            m_Key = value
        End Set
    End Property
    Private m_Key As String
    Public Property Action() As String
        Get
            Return m_Action
        End Get
        Set(value As String)
            m_Action = value
        End Set
    End Property
    Private m_Action As String
End Class

Public Class ControlsManager
    Inherits Collection(Of ControlsItems)

    Public Event Changed As EventHandler(Of ControlsChangedEventArgs)

    Protected Overrides Sub InsertItem( _
        ByVal index As Integer, ByVal newItem As ControlsItems)

        MyBase.InsertItem(index, newItem)

        RaiseEvent Changed(Me, New ControlsChangedEventArgs( _
            ChangeType.Added, newItem, Nothing))
    End Sub

    Protected Overrides Sub SetItem(ByVal index As Integer, _
        ByVal newItem As ControlsItems)

        Dim replaced As ControlsItems = Items(index)
        MyBase.SetItem(index, newItem)

        RaiseEvent Changed(Me, New ControlsChangedEventArgs( _
            ChangeType.Replaced, replaced, newItem))
    End Sub

    Protected Overrides Sub RemoveItem(ByVal index As Integer)

        Dim removedItem As ControlsItems = Items(index)
        MyBase.RemoveItem(index)

        RaiseEvent Changed(Me, New ControlsChangedEventArgs( _
            ChangeType.Removed, removedItem, Nothing))
    End Sub

    Protected Overrides Sub ClearItems()
        MyBase.ClearItems()

        RaiseEvent Changed(Me, New ControlsChangedEventArgs( _
            ChangeType.Cleared, Nothing, Nothing))
    End Sub

End Class

' Event argument for the Changed event. 
' 
Public Class ControlsChangedEventArgs
    Inherits EventArgs

    Public ReadOnly ChangedItem As ControlsItems
    Public ReadOnly ChangeType As ChangeType
    Public ReadOnly ReplacedWith As ControlsItems

    Public Sub New(ByVal change As ChangeType, ByVal item As ControlsItems, _
        ByVal replacement As ControlsItems)

        ChangeType = change
        ChangedItem = item
        ReplacedWith = replacement
    End Sub
End Class

Public Enum ChangeType
    Added
    Removed
    Replaced
    Cleared
End Enum