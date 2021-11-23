Imports System.ComponentModel
Imports System.IO
Imports System.Xml.Serialization
#Region "����������"
''' <summary>
''' ��������� ��� ������ � ����������� � ����������
''' </summary>
''' <remarks>������ �������� ����� ��� ��������� � ������ ������ ��������� �� ������� �������</remarks>
Public Interface IMessageWorker
    ''' <summary>
    ''' �������������� ����� ���������
    ''' </summary>
    ''' <param name="msgText">����� ���������</param>
    ''' <param name="msgTitle">��������� ���������</param>
    ''' <param name="msgOptons">����� ���������</param>
    Function ShowMessage(msgText As String, Optional msgTitle As String = "��������!", Optional msgOptons As MessageOptions = Nothing) As Task(Of Boolean)
End Interface
#End Region
#Region "������������� � ����"
''' <summary>
''' ����� ���������
''' </summary>
Public Enum MessageStyle
    ''' <summary>
    ''' ��������� �� ������
    ''' </summary>
    ErrorMessage = 0
    ''' <summary>
    ''' �������������� ���������
    ''' </summary>
    InfoMessage = 1
    ''' <summary>
    ''' ��������� ��������� ������ "��" ��� "���"
    ''' </summary>
    YesNo = 2
End Enum
''' <summary>
''' �������� ��������� ������� � API
''' </summary>
Public Structure APIQeryResult
    Public Sub New(sr As Boolean, rb As Object)
        IsSuccessfulRequest = sr
        ResulBody = rb
    End Sub
    ''' <summary>
    ''' True ���� ������ ��������
    ''' </summary>
    Dim IsSuccessfulRequest As Boolean
    ''' <summary>
    ''' ������ ���������� �������
    ''' </summary>
    Dim ResulBody As Object
End Structure
''' <summary>
''' ��������� ���� ��������� �� ������
''' </summary>
Public Class MessageOptions
    ''' <summary>
    ''' ����� ���������
    ''' </summary>
    Public MessageStyle As MessageStyle = MessageStyle.InfoMessage
    ''' <summary>
    ''' ���������� �� ��������� ������ ��������� �������� ����
    ''' </summary>
    Public IsTopMost As Boolean = False
    ''' <summary>
    ''' �������� �� ��������� ������������� �� ��������� ��������� ����� ������
    ''' </summary>
    Public IsAutoHide As Boolean = True
End Class
#End Region
#Region "������� ������"
''' <summary>
''' ������� ����� ��� ���� �������, ��� ���������� �������� ������
''' </summary>
''' <remarks>��������� ���������� ���� � �������, ������� ��������� �������� ������</remarks>
Public MustInherit Class NotifyProperty_Base(Of T)
    Implements INotifyPropertyChanged
#Region "���������� ���������� INotifyPropertyChanged"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Shadows Sub OnPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub
#End Region
    Public Overridable Sub SetPropertys(input As T, Optional ignoreBaseType As String = "")
        Dim inputType As Type = input.GetType
        Dim meType As Type = [GetType]()
        '�������� ������ �� ��������� �������� ������
        For Each inputProperty In inputType.GetProperties
            Dim meProperty = meType.GetProperty(inputProperty.Name)
            If meProperty IsNot Nothing Then
                If ignoreBaseType = "" OrElse meProperty.PropertyType.BaseType.ToString.IndexOf(ignoreBaseType) = -1 Then
                    If meProperty.PropertyType.BaseType.ToString.IndexOf("NotifyProperty_Base") > -1 Then
                        meProperty.PropertyType.GetMethod("SetPropertys").Invoke(meProperty.GetValue(Me), {inputProperty.GetValue(input), ""})
                    Else
                        meProperty.SetValue(Me, inputProperty.GetValue(input))
                    End If
                End If
            End If
        Next
    End Sub
End Class

''' <summary>
''' ������� ����� ��� ������ � ������� ����������.
''' </summary>
''' <remarks>��������� ������� ��������, ������� ������ ���� � ������ ���� ����������</remarks>
Public MustInherit Class ApplicationData_Base
    Inherits NotifyProperty_Base(Of ApplicationData_Base)
#Region "������ � ����������� ��������"
    ''' <summary>
    ''' ���������������� ��������, ������� ��������� �������� ����� � ����������� ����������
    ''' </summary>
    ''' <returns>���������� ������ � ��������� �����</returns>
    ''' <remarks>������ ��� ���������� ������ �������� ��� �����, ���� ����������� ��� ���������</remarks>
    Protected MustOverride Property AppSettingPath As String

    ''' <summary>
    ''' �������, ������� ��������� ������ ������ ���� � ����� �������� ����������
    ''' </summary>
    ''' <remarks>���� AppSettingPath �� �����, �� ����� ����������� LXGApplications</remarks>
    Private Function GetSavePath() As String
        Return IO.Path.Combine(IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), If(AppSettingPath, "LXGApplications")), "appd.dat")
    End Function

    ''' <summary>
    ''' ��������� ������ ���������� � ����
    ''' </summary>
    ''' <param name="s">����� � �������</param>
    Public Sub SaveSettings(Of T)(s As T)
        Try
            If Not IO.File.Exists(GetSavePath) Then
                Directory.CreateDirectory(Path.GetDirectoryName(GetSavePath))
            End If
            Dim writer As New XmlSerializer(s.GetType)
            Dim file As New StreamWriter(GetSavePath)
            writer.Serialize(file, s)
            file.Close()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' ��������� ������ ���������� �� �����
    ''' </summary>
    ''' <typeparam name="T">��� ������ � ������� ����������� ������</typeparam>
    Public Function LoadSettings(Of T As {Class, New})() As T
        If File.Exists(GetSavePath) Then
            Try
                Dim reader = New XmlSerializer(GetType(T))
                Dim file = New StreamReader(GetSavePath)
                Dim result As T = reader.Deserialize(file)
                file.Close()
                Return result
            Catch ex As Exception
                Return New T
            End Try
        Else
            Return New T
        End If
    End Function
    ''' <summary>
    ''' ������� ���� ��������, ��� �������� � ������ ���������� � ���������� �� ��������� (����� �����������)
    ''' </summary>
    Public Sub SetDefaultSettings()
        If File.Exists(GetSavePath) Then
            File.Delete(GetSavePath)
        End If
    End Sub
#End Region

#Region "������ � ������������� (�����������) ����������"
    ''' <summary>
    ''' ���������� ����� ��� ������ � �����������, ������������� � �������������� ������
    ''' </summary>
    ''' <remarks>� �������������� ApplicationData ������ ����� ������� ��������� ������ ������������ IMessageWorker � �������� ���� ��������� ����� ��� �������</remarks>
    Protected MustOverride Function GetMessageWorker() As IMessageWorker
#End Region

#Region "������ � API � ������ ������"
    ''' <summary>
    ''' ���������� � API ��� ���� ������ ����� ������� �� ������� ��������� func
    ''' </summary>
    ''' <typeparam name="T">��� ������ ������������� ��������</typeparam>
    ''' <param name="func">������ �� ������� ����������� ���������������� ������</param>
    ''' <param name="args">������ ���������� ������������ �������</param>
    ''' <param name="msgOptions">��������� ���� ��������� �� ������</param>
    ''' <remarks>
    ''' ������� �������, ������� ������������� ��������� �� ������ ������� � API ��� ���� ������.
    ''' � ��������� ���������� ������� ��������� �� ������ ��� ���������� �� �������
    ''' </remarks>
    Public Async Function APIQery(Of T)(func As Func(Of Object, Task(Of APIQeryResult)), args As Object(), Optional msgOptions As MessageOptions = Nothing) As Task(Of T)
        Dim Result As APIQeryResult = Await Await Task.Factory.StartNew(func, args)
        If Result.IsSuccessfulRequest Then
            Return Result.ResulBody
        Else
            If Result.ResulBody Is Nothing Then
                Await GetMessageWorker.ShowMessage("��� ��������� � ������� ��������� ������!",, msgOptions)
            Else
                Await GetMessageWorker.ShowMessage(Uri.UnescapeDataString(Result.ResulBody),, msgOptions)
            End If
            Return Nothing
        End If
    End Function
#End Region
End Class
''' <summary>
''' ������� ����� ��� ������ �������� �����
''' </summary>
Public MustInherit Class CloudWorker_Base
    ''' <summary>
    ''' ���������� ����� ��� ������ � �����������, ������������� � �������������� ������
    ''' </summary>
    ''' <remarks>� �������������� ApplicationData ������ ����� ������� ��������� ������ ������������ IMessageWorker � �������� ���� ��������� ����� ��� �������</remarks>
    Protected MustOverride Function GetMessageWorker() As IMessageWorker
    ''' <summary>
    ''' ���������� � API ��� ���� ������ ����� ������� �� ������� ��������� func
    ''' </summary>
    ''' <typeparam name="T">��� ������ ������������� ��������</typeparam>
    ''' <param name="func">������ �� ������� ����������� ���������������� ������</param>
    ''' <param name="args">������ ���������� ������������ �������</param>
    ''' <remarks>
    ''' ������� �������, ������� ������������� ��������� �� ������ ������� � API ��� ���� ������.
    ''' � ��������� ���������� ������� ��������� �� ������ ��� ���������� �� �������
    ''' </remarks>
    Protected Async Function APIQery(Of T)(func As Func(Of Object, Task(Of APIQeryResult)), args As Object()) As Task(Of T)
        Dim Result As APIQeryResult = Await Await Task.Factory.StartNew(func, args)
        If Result.IsSuccessfulRequest Then
            Return Result.ResulBody
        Else
            If Result.ResulBody Is Nothing Then
                Await GetMessageWorker.ShowMessage("��� ��������� � ������� ��������� ������!",, Nothing)
            Else
                Await GetMessageWorker.ShowMessage(Uri.UnescapeDataString(Result.ResulBody),, Nothing)
            End If
            Return Nothing
        End If
    End Function
End Class
#End Region