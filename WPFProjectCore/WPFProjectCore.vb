Imports System.ComponentModel
Imports System.IO
Imports System.Xml.Serialization
#Region "Интерфейсы"
''' <summary>
''' Интерфейс для работы с сообщениями в приложении
''' </summary>
''' <remarks>Данный интефейс нужен для обращения к методу показа сообщений из базовых классов</remarks>
Public Interface IMessageWorker
    ''' <summary>
    ''' Инициализирует показ сообщения
    ''' </summary>
    ''' <param name="msgText">Текст сообщения</param>
    ''' <param name="msgTitle">Заголовок сообщения</param>
    ''' <param name="msgOptons">Опции сообщения</param>
    Function ShowMessage(msgText As String, Optional msgTitle As String = "Внимание!", Optional msgOptons As MessageOptions = Nothing) As Task(Of Boolean)
End Interface
#End Region
#Region "Перечеслители и типы"
''' <summary>
''' Стиль сообщения
''' </summary>
Public Enum MessageStyle
    ''' <summary>
    ''' Сообщение об ошибке
    ''' </summary>
    ErrorMessage = 0
    ''' <summary>
    ''' Информационное сообщение
    ''' </summary>
    InfoMessage = 1
    ''' <summary>
    ''' Сообщение требующие ответа "Да" или "Нет"
    ''' </summary>
    YesNo = 2
End Enum
''' <summary>
''' Содержит результат запроса к API
''' </summary>
Public Structure APIQeryResult
    Public Sub New(sr As Boolean, rb As Object)
        IsSuccessfulRequest = sr
        ResulBody = rb
    End Sub
    ''' <summary>
    ''' True если запрос успешный
    ''' </summary>
    Dim IsSuccessfulRequest As Boolean
    ''' <summary>
    ''' Объект результата запроса
    ''' </summary>
    Dim ResulBody As Object
End Structure
''' <summary>
''' Настройки окна сообщения об ошибке
''' </summary>
Public Class MessageOptions
    ''' <summary>
    ''' Стиль сообщения
    ''' </summary>
    Public MessageStyle As MessageStyle = MessageStyle.InfoMessage
    ''' <summary>
    ''' Показываль ли сообщения поверх основного контента окна
    ''' </summary>
    Public IsTopMost As Boolean = False
    ''' <summary>
    ''' Скрывать ли сообщение автоматически по истечении заданного числа секунд
    ''' </summary>
    Public IsAutoHide As Boolean = True
End Class
#End Region
#Region "Базовые классы"
''' <summary>
''' Базовый класс для всех классов, что используют привязку данных
''' </summary>
''' <remarks>Уменьшает количество кода в классах, которые реализуют привязку данных</remarks>
Public MustInherit Class NotifyProperty_Base(Of T)
    Implements INotifyPropertyChanged
#Region "Реализация интерфейса INotifyPropertyChanged"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Shadows Sub OnPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub
#End Region
    Public Overridable Sub SetPropertys(input As T, Optional ignoreBaseType As String = "")
        Dim inputType As Type = input.GetType
        Dim meType As Type = [GetType]()
        'Проходим циклом по свойствам текущего класса
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
''' Базовый класс для работы с данными приложения.
''' </summary>
''' <remarks>Реализует базовые механики, которые должны быть в каждом моем приложении</remarks>
Public MustInherit Class ApplicationData_Base
    Inherits NotifyProperty_Base(Of ApplicationData_Base)
#Region "Работа с сохранением настроек"
    ''' <summary>
    ''' Переопределяемое свойство, которое указаыает название папки с настройками приложения
    ''' </summary>
    ''' <returns>Возвращает строку с названием папки</returns>
    ''' <remarks>Каждое мое приложение должно задавать имя папки, куда сохраняются его настройки</remarks>
    Protected MustOverride Property AppSettingPath As String

    ''' <summary>
    ''' Функция, которая формирует полную строку пути к файлу настроек приложения
    ''' </summary>
    ''' <remarks>Если AppSettingPath не задан, то папка поумолчанию LXGApplications</remarks>
    Private Function GetSavePath() As String
        Return IO.Path.Combine(IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), If(AppSettingPath, "LXGApplications")), "appd.dat")
    End Function

    ''' <summary>
    ''' Сохраняет данные приложения в файл
    ''' </summary>
    ''' <param name="s">Класс с данными</param>
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
    ''' Считывает данные приложения из файла
    ''' </summary>
    ''' <typeparam name="T">Тип класса в который считываются данные</typeparam>
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
    ''' Удаляет файл настроек, что приводит к сбросу приложения к настройкам по умолчанию (после перезапуска)
    ''' </summary>
    Public Sub SetDefaultSettings()
        If File.Exists(GetSavePath) Then
            File.Delete(GetSavePath)
        End If
    End Sub
#End Region

#Region "Работа с уведомлениями (сообщениями) приложения"
    ''' <summary>
    ''' Возвращает класс для работы с сообщениями, реализованный в унаследованном классе
    ''' </summary>
    ''' <remarks>В унаследованном ApplicationData классе нужно вернуть экземпляр класса реализующего IMessageWorker к которому идет обращение через эту функцию</remarks>
    Protected MustOverride Function GetMessageWorker() As IMessageWorker
#End Region

#Region "Работа с API и базами данных"
    ''' <summary>
    ''' Обращается к API или базе данных через функцию на которую ссылается func
    ''' </summary>
    ''' <typeparam name="T">Тип класса возвращаемого функцией</typeparam>
    ''' <param name="func">Ссылка на функцию выполняющую непосредственный запрос</param>
    ''' <param name="args">Массив аргументов передаваемых функции</param>
    ''' <param name="msgOptions">Настройки окна сообщения об ошибке</param>
    ''' <remarks>
    ''' Базовая функция, которая автоматически реагирует на ошибку запроса к API или базе данных.
    ''' В частности показывает базовое сообщение об ошибке или переданное из функции
    ''' </remarks>
    Public Async Function APIQery(Of T)(func As Func(Of Object, Task(Of APIQeryResult)), args As Object(), Optional msgOptions As MessageOptions = Nothing) As Task(Of T)
        Dim Result As APIQeryResult = Await Await Task.Factory.StartNew(func, args)
        If Result.IsSuccessfulRequest Then
            Return Result.ResulBody
        Else
            If Result.ResulBody Is Nothing Then
                Await GetMessageWorker.ShowMessage("При обращении к серверу произашла ошибка!",, msgOptions)
            Else
                Await GetMessageWorker.ShowMessage(Uri.UnescapeDataString(Result.ResulBody),, msgOptions)
            End If
            Return Nothing
        End If
    End Function
#End Region
End Class
''' <summary>
''' Базовый класс для работы облачных служб
''' </summary>
Public MustInherit Class CloudWorker_Base
    ''' <summary>
    ''' Возвращает класс для работы с сообщениями, реализованный в унаследованном классе
    ''' </summary>
    ''' <remarks>В унаследованном ApplicationData классе нужно вернуть экземпляр класса реализующего IMessageWorker к которому идет обращение через эту функцию</remarks>
    Protected MustOverride Function GetMessageWorker() As IMessageWorker
    ''' <summary>
    ''' Обращается к API или базе данных через функцию на которую ссылается func
    ''' </summary>
    ''' <typeparam name="T">Тип класса возвращаемого функцией</typeparam>
    ''' <param name="func">Ссылка на функцию выполняющую непосредственный запрос</param>
    ''' <param name="args">Массив аргументов передаваемых функции</param>
    ''' <remarks>
    ''' Базовая функция, которая автоматически реагирует на ошибку запроса к API или базе данных.
    ''' В частности показывает базовое сообщение об ошибке или переданное из функции
    ''' </remarks>
    Protected Async Function APIQery(Of T)(func As Func(Of Object, Task(Of APIQeryResult)), args As Object()) As Task(Of T)
        Dim Result As APIQeryResult = Await Await Task.Factory.StartNew(func, args)
        If Result.IsSuccessfulRequest Then
            Return Result.ResulBody
        Else
            If Result.ResulBody Is Nothing Then
                Await GetMessageWorker.ShowMessage("При обращении к серверу произашла ошибка!",, Nothing)
            Else
                Await GetMessageWorker.ShowMessage(Uri.UnescapeDataString(Result.ResulBody),, Nothing)
            End If
            Return Nothing
        End If
    End Function
End Class
#End Region