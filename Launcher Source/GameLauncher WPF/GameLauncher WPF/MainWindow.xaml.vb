Imports System.Diagnostics
Imports System.IO
Imports System.Net

#Region "Window"

Class MainWindow

    '#Region "Properties"

    '    Private tmpKey As String
    '    Private m_key As String
    '    Private Property KeyBoardHook
    '        Get
    '            Return m_key
    '        End Get
    '        Set(value)
    '            m_key = value
    '        End Set
    '    End Property

    '#End Region

#Region "Variable declarations"

#Region "Writable variables"

#Region "Shared variables"

    Public Shared IsGameInstalled As Boolean = True
    Public Shared TutoShown As Boolean = False

    Public Shared AppWorkPath As String = AppData & "\.dynawars\"
    Public Shared GamePath As String = Paths.GamePath

#End Region

#Region "Public variables"

    Public Maxfps_all As String = ""
    Public Brightness_all As String = ""
    Public FOV_all As String = ""
    Public Sens_All As String = ""
    Public wbWindow As WebBrowserOverlay

#End Region

#Region "Local variables"

    Private pInfo As ProcessStartInfo = New ProcessStartInfo
    Private WithEvents Proceso As New Process

    Private wc As WebClient
    Private sw As Stopwatch = New Stopwatch

    Dim GameUrl As String = "" 'Sets the Game Url .zip
    Dim DownloadPath As String = "" 'The path where the game will be downloaded (temp folder)
    Dim NeedToExtract As Boolean = True 'This will say if the downloaded file needs to be extracted
    Dim PathToExtract As String = "" 'This path is the path where the game will be extracted

    Dim IsAppRunningOnEditor As Boolean = False

    Dim WithEvents winWebBrowser As WebBrowser

#End Region

#End Region

#Region "ReadOnly Variable Declarations"

#Region "Shared variables"

    Public Shared ReadOnly AppData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Public Shared ReadOnly OutputFile As String = Path.GetTempFileName
    Public Shared ReadOnly GameKey = "1"

#End Region

#Region "Local Varaibles"



#End Region

#End Region

#End Region

#Region "Local Functions"

    Public Sub DownloadFile(urlAddress As String, location As String)
        Using webClient = New WebClient()

            AddHandler wc.DownloadProgressChanged, AddressOf ProgressChanged
            AddHandler wc.DownloadFileCompleted, AddressOf Completed

            ' The variable that will be holding the url address (making sure it starts with http://)
            Dim URL As Uri = If(urlAddress.StartsWith("http://", StringComparison.OrdinalIgnoreCase), New Uri(urlAddress), New Uri("http://" & urlAddress))

            ' Start the stopwatch which we will be using to calculate the download speed
            sw.Start()

            Try
                ' Start downloading the file
                webClient.DownloadFileAsync(URL, location)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using
    End Sub

    ' The event that will fire whenever the progress of the WebClient is changed
    Private Sub ProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)

        ' Show all stats
        lblProgress.Content = String.Format("Descargando archivos... {0}; Restante: {1}; TEL: {2}; Porcentaje: {3}; Tamaño total del paquete: {4}", _
                                            SpecialTextFormat.FormatBytes(e.BytesReceived / 1024.0 / sw.Elapsed.TotalSeconds) & "\s", _
                                            SpecialTextFormat.FormatBytes((e.TotalBytesToReceive / 1024.0 / 1024.0) - (e.BytesReceived / 1024.0 / 1024.0)), _
                                            SpecialTextFormat.FormatSeconds(e.TotalBytesToReceive / (e.BytesReceived / sw.Elapsed.TotalSeconds)), e.ProgressPercentage.ToString() & "%", _
                                            SpecialTextFormat.FormatBytes(e.TotalBytesToReceive / 1024.0 / 1024.0))

        ' Update the progressbar percentage only when the value is not the same.
        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    ' The event that will trigger when the WebClient is completed
    Private Sub Completed(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs)
        ' Reset the stopwatch.
        sw.Reset()

        If NeedToExtract Then
            'Extract the zip
            ZipManager.Extract(DownloadPath, PathToExtract)
        End If

        If e.Cancelled = True Then
            lblProgress.Content = "La descarga fue cancelada."
        Else
            lblProgress.Content = "La descarga ha sido completada."
        End If
    End Sub

    Private Shadows Sub TSLoaded()
        TreeHelper.FindVisualChildByName(Of Image)(itemControls, "iconKB").Source = New BitmapImage(New Uri("pack://application:,,,/Resources/icons/Keyboard.png", UriKind.RelativeOrAbsolute))
        TreeHelper.FindVisualChildByName(Of Image)(itemMouse, "iconMouse").Source = New BitmapImage(New Uri("pack://application:,,,/Resources/icons/Mouse.png", UriKind.RelativeOrAbsolute))
        TreeHelper.FindVisualChildByName(Of Image)(itemAudio, "iconAudio").Source = New BitmapImage(New Uri("pack://application:,,,/Resources/icons/Audio.png", UriKind.RelativeOrAbsolute))
        TreeHelper.FindVisualChildByName(Of Image)(itemVideo, "iconVideo").Source = New BitmapImage(New Uri("pack://application:,,,/Resources/icons/Video.png", UriKind.RelativeOrAbsolute))
        TreeHelper.FindVisualChildByName(Of Image)(itemSettings, "iconSettings").Source = New BitmapImage(New Uri("pack://application:,,,/Resources/icons/Settings.png", UriKind.RelativeOrAbsolute))
    End Sub

#End Region

#Region "Event functions"

    Private Shadows Sub Load() Handles MyBase.Loaded

        'Load the WebBrowser

        wbWindow = New WebBrowserOverlay(_webBrowserPlacementTarget)

        winWebBrowser = wbWindow.WebBrowser

        'Load another variables

        IsAppRunningOnEditor = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.Contains(".vshost.exe")

        IsGameInstalled = Directory.Exists(GamePath & "Game_Data")

        AppWorkPath = IOUtils.ValidatePath(AppWorkPath)
        GamePath = IOUtils.ValidatePath(GamePath)

        Launcher.CreateFolders()

        If IsGameInstalled And Not String.IsNullOrEmpty(Profile.UserName) Then
            Launcher.LoadAllProfiles()
        End If

        If cmbProfiles.Items.Count > 0 Then
            cmbProfiles.Text = cmbProfiles.Items(0).ToString
        End If

        txtBoxW.Foreground = Brushes.LightGray
        txtBoxH.Foreground = Brushes.LightGray

        Profile.IsLoggedIn = False

        loginPanel.Visibility = Windows.Visibility.Visible

        'Load all the images

        Fondo.Source = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/fondo.png", UriKind.RelativeOrAbsolute))
        Logo.Source = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/Logo.png", UriKind.RelativeOrAbsolute))
        imgCompleted.Source = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/tick-icon.png", UriKind.RelativeOrAbsolute))

        Dim PbWheel As ImageSource = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/progressbar.gif", UriKind.RelativeOrAbsolute))
        WpfAnimatedGif.ImageBehavior.SetAnimatedSource(animProgress, PbWheel)
        WpfAnimatedGif.ImageBehavior.SetRepeatBehavior(animProgress, Animation.RepeatBehavior.Forever)

        'Checks if the game is downloaded if not download it

        If Not IsGameInstalled And Not GamePath.Equals(String.Empty) And Not IsAppRunningOnEditor Then

            'Set all the variables
            GameUrl = ""
            DownloadPath = AppData & "\.dynawars\temp\Game.zip"
            NeedToExtract = True
            PathToExtract = GamePath

            DownloadFile(GameUrl, DownloadPath)

        End If

        'Checks if the game needs to update

        Dim FixedGameKeyURL = "https://dl.dropboxusercontent.com/s/s9exi4dkftg0jls/Dynawars%20Game%20Key%20Updater.txt?dl=1"

        Dim DownloadURL As String = Updater.CheckAndGet(FixedGameKeyURL, GameKey)

        Dim IsAviableUpdate As Boolean = Not String.IsNullOrEmpty(DownloadURL)

        If IsAviableUpdate Then

            If MsgBox("Hay una actualizacion disponible para el juego, ¿deseas descargarla?", MsgBoxStyle.YesNo, "Información") = MsgBoxResult.Yes Then

                DownloadPath = AppData & "\.dynawars\temp\Game.zip"
                NeedToExtract = True
                PathToExtract = IOUtils.NewGamePath("Introduzca la ruta de instalación")

                DownloadFile(DownloadURL, DownloadPath)

            End If

        End If

        If newFeatures.IsSelected Then

            wbWindow.Visibility = Windows.Visibility.Visible

        End If

        'Set profiles columns

        Dim c1 As New DataGridTextColumn()
        c1.Header = "Nombre"
        c1.Binding = New Binding("ProfileName")
        c1.IsReadOnly = True
        c1.Width = 241
        dtProfiles.Columns.Add(c1)

        Dim c2 As New DataGridTextColumn()
        c2.Header = "Versión"
        c2.Binding = New Binding("VersionName")
        c2.IsReadOnly = True
        c2.Width = 241
        dtProfiles.Columns.Add(c2)

        Dim c3 As New DataGridTextColumn()
        c3.Header = "Usuario"
        c3.Binding = New Binding("Username")
        c3.IsReadOnly = True
        c3.Width = 241
        dtProfiles.Columns.Add(c3)

        'Add an empty row

        'dtProfiles.Items.Add(New ProfileItem() With { _
        '        .ProfileName = "", _
        '        .VersionName = "", _
        '        .Username = "" _
        '    })

        'Add controls columns

        Dim c4 As New DataGridTextColumn()
        c4.Header = "Acción"
        c4.Binding = New Binding("Action")
        c4.IsReadOnly = True
        c4.Width = 241
        dtControls.Columns.Add(c4)

        Dim c5 As New DataGridTextColumn()
        c5.Header = "Tecla/Botón"
        c5.Binding = New Binding("Key")
        c5.IsReadOnly = True
        c5.Width = 241
        dtControls.Columns.Add(c5)

        'Add an empty row

        'dtControls.Items.Add(New ControlsItems() With { _
        '    .Action = "", _
        '    .Key = "" _
        '})

        If String.IsNullOrEmpty(user.Text) Then
            login.IsEnabled = False
        End If

        'Add handlers for check that

        AddHandler user.TextChanged, AddressOf LoginButtonDisable
        AddHandler pass.PasswordChanged, AddressOf LoginButtonDisable

        'Write label

        lblProgress.Content = "Todo en orden."

    End Sub

    Protected Overrides Sub OnContentRendered(e As EventArgs)
        MyBase.OnContentRendered(e)

        'ShowTuto()

    End Sub

    Protected Overrides Sub OnMouseLeftButtonDown(ByVal e As MouseButtonEventArgs)

        MyBase.OnMouseLeftButtonDown(e)
        Me.DragMove()

    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        End
    End Sub

    Private Sub Button_Click_2(sender As Object, e As RoutedEventArgs) Handles btnPlay.Click

        If Not File.Exists(GamePath & "Game.exe") Then
            MsgBox("El juego no se encuentra en donde debería estar, o tocaste algo de la configuración." & Environment.NewLine & "Por favor, para volver a la normalidad, borra el archivo " & AppData & "\.dynawars\app.conf, cierra el Launcher y vuelve a abrirlo.", MsgBoxStyle.Critical, "Error")
        End If

        If cmbProfiles.Items.Count = 0 Then
            MsgBox("No has creado todavía ningún perfil debes crear uno y seleccionarlo para poder jugar.", MsgBoxStyle.Exclamation, "Advertencia")
            Exit Sub
        End If

        If cmbProfiles.SelectedIndex = -1 Then
            MsgBox("Debes seleccionar un perfil para poder jugar.", MsgBoxStyle.Exclamation, "Advertencia")
            Exit Sub
        End If

        Dim width As String
        Dim height As String
        Dim fullscreen As String

        'If Not File.Exists(AppWorkPath & "Launcher\Settings\" & "launcherConfig.cfg") Then
        '    MsgBox("Nunca has editado las 'preferencias del Launcher' debes editarlas y guardarlas para poder jugar.", MsgBoxStyle.Exclamation, "Advertencia")
        '    Exit Sub
        'Else
        '    If Not txtBoxW.Text.Equals(INIFileManager.Key.Get("app-Width")) And Not txtBoxW.Text.Equals("Anchura") Or Not txtBoxH.Text.Equals(INIFileManager.Key.Get("app-Height")) And Not txtBoxH.Text.Equals("Altura") Then
        '        MsgBox("No has guardado las 'preferencias del Launcher', debes guardarlo para poder jugar.", MsgBoxStyle.Exclamation, "Advertencia")
        '        Exit Sub
        '    End If
        'End If

        If String.IsNullOrEmpty(Profile.LoadedProfile) Then
            MsgBox("Ningún perfil ha sido cargado, por favor, vuelve a seleccionar uno para poder jugar.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If

        INIFileManager.FilePath = Path.Combine(MainWindow.GamePath & "Game_Data\Profiles\" & Profile.LoadedProfile, "\profile.cfg")

        width = INIFileManager.Key.Get("app-Width")
        height = INIFileManager.Key.Get("app-Height")
        fullscreen = INIFileManager.Key.Get("fullScreen")

        'If Not File.Exists(GamePath & "Game.exe") Then
        '    MsgBox("Has movido el Launcher de lugar, ponlo donde estaba antes para poder jugar...", MsgBoxStyle.Exclamation, "Advertencia")
        '    Exit Sub
        'End If

        With pInfo
            .FileName = GamePath & "Game.exe"
            .Arguments = String.Format("-screen-width {0} -screen-height {1} -CustomArgs:fullscreen={2};profile={3}", width, height, fullscreen, Profile.LoadedProfile)
            .CreateNoWindow = True
            .UseShellExecute = False
        End With

        With Proceso
            .EnableRaisingEvents = True
            .StartInfo = pInfo
        End With

        Proceso.Start()

    End Sub

    Private Sub Button_Click_3(sender As Object, e As RoutedEventArgs)
        If Not Profile_Name.Text.Equals(String.Empty) Then
            If Not String.IsNullOrEmpty(Profile.UserName) Then
                Launcher.SaveProfile()
            Else
                MsgBox("Debes loguearte antes de poder guardar cualquier perfil.", MsgBoxStyle.Critical, "Error")
            End If
        Else
            MsgBox("Debe poner un nombre a este Perfil antes de 'Guardar' nada.", MsgBoxStyle.Critical, "Error")
        End If
    End Sub

    'Private Sub ProgressBar_Loaded_1(sender As Object, e As RoutedEventArgs) Handles ProgressBar1.Loaded
    '    Dim p = DirectCast(sender, ProgressBar)
    '    p.ApplyTemplate()

    '    'FW 4 problems
    '    DirectCast(p.Template.FindName("Animation", p), Panel).Background = Brushes.Transparent
    'End Sub

    Private Sub ProgressBar1_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ProgressBar1.ValueChanged
        If ProgressBar1.Value = ProgressBar1.Maximum Then
            img.ImageSource = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/pbc.png", UriKind.RelativeOrAbsolute))
            animProgress.Visibility = Windows.Visibility.Hidden
            imgCompleted.Visibility = Windows.Visibility.Visible
            btnPlay.IsEnabled = True
            btnPlay.Foreground = Brushes.Black
        Else
            img.ImageSource = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/pbf.png", UriKind.RelativeOrAbsolute))
            animProgress.Visibility = Windows.Visibility.Visible
            imgCompleted.Visibility = Windows.Visibility.Hidden
            btnPlay.IsEnabled = False
            btnPlay.Foreground = Brushes.Gray
        End If
    End Sub

    Private Sub ComboResW_DClick(sender As Object, e As MouseButtonEventArgs)
        ComboBoxW.Visibility = Windows.Visibility.Hidden
        'ComboBoxH.Visibility = Windows.Visibility.Hidden
        'txtBoxH.Visibility = Windows.Visibility.Visible
        txtBoxW.Visibility = Windows.Visibility.Visible
        brdWidth.Visibility = Windows.Visibility.Visible
        'brdHeight.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub ComboResH_DClick(sender As Object, e As MouseButtonEventArgs)
        'ComboBoxW.Visibility = Windows.Visibility.Hidden
        ComboBoxH.Visibility = Windows.Visibility.Hidden
        txtBoxH.Visibility = Windows.Visibility.Visible
        'txtBoxW.Visibility = Windows.Visibility.Visible
        'brdWidth.Visibility = Windows.Visibility.Visible
        brdHeight.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub txtBoxW_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtBoxW.GotFocus
        If txtBoxW.Text = "Anchura" Then
            txtBoxW.Foreground = Brushes.Black
            txtBoxW.Text = ""
        End If
    End Sub

    Private Sub txtBoxW_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtBoxW.LostFocus
        If txtBoxW.Text.Length = 0 Then
            txtBoxW.Foreground = Brushes.LightGray
            txtBoxW.Text = "Anchura"
            ComboBoxW.Visibility = Windows.Visibility.Visible
        Else
            ComboBoxW.Visibility = Windows.Visibility.Visible
        End If
    End Sub

    Private Sub txtBoxH_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtBoxH.GotFocus
        If txtBoxH.Text = "Altura" Then
            txtBoxH.Foreground = Brushes.Black
            txtBoxH.Text = ""
        End If
    End Sub

    Private Sub txtBoxH_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtBoxH.LostFocus
        If txtBoxH.Text.Length = 0 Then
            txtBoxH.Foreground = Brushes.LightGray
            txtBoxH.Text = "Altura"
            ComboBoxH.Visibility = Windows.Visibility.Visible
        Else
            ComboBoxH.Visibility = Windows.Visibility.Visible
        End If
    End Sub

    Private Sub tabSettings_Click(sender As Object, e As MouseButtonEventArgs) 'Esto no me acuerdo porque lo hice
        'ComboBoxW.Visibility = Windows.Visibility.Visible
        'ComboBoxH.Visibility = Windows.Visibility.Visible

        'txtBoxH.Text = ""
        'txtBoxW.Text = ""
    End Sub

    Private Sub Sensibility_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double))

        Dim sens_res As Double = ((Sensibility.Value - 0.1) / (30 - 0.1)) * 200

        If Sensibility.Value = 30 Then
            Sens_All = "HIPERVELOCIDAD!"
        ElseIf Sensibility.Value = 0.1 Then
            Sens_All = "*bostezo*"
        Else
            Sens_All = sens_res.ToString("F0")
        End If

        lblSens.Content = "Sensibilidad [" & Sens_All & "]:"

    End Sub

    Private Sub Button_Click_4(sender As Object, e As RoutedEventArgs)

        If cmbProfiles.Items.Count = 0 Then
            MsgBox("No has creado todavía ningún perfil debes crear uno y seleccionarlo para poder editarlo.", MsgBoxStyle.Exclamation, "Advertencia")
            Exit Sub
        End If

        tabCSettings.SelectedIndex = 2

        If Not String.IsNullOrEmpty(Profile.UserName) Then
            Launcher.LoadProfile(DirectCast(ComboBoxH.SelectedItem, ComboBoxItem).Content.ToString())
        End If

    End Sub

    Private Sub FOV_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles FOV.ValueChanged
        Dim fov_res As Double = ((FOV.Value - 60) / (120 - 60)) * 200

        If FOV.Value = 60 Then
            FOV_all = "Normal"
        ElseIf FOV.Value = 120 Then
            FOV_all = "Quake Pro"
        Else
            FOV_all = fov_res.ToString("F0")
        End If

        lblFOV.Content = "FOV (Field Of View) [" & FOV_all & "]:"
    End Sub

    Private Sub RenderDistance_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles RenderDistance.ValueChanged
        lblRenderDistance.Content = "Distancia de renderizado [" & (RenderDistance.Value / 1000).ToString("F1") & " Km]:"
    End Sub

    Private Sub Button_Click_5(sender As Object, e As RoutedEventArgs)

        tabCSettings.SelectedIndex = 2

        If Not String.IsNullOrEmpty(Profile.UserName) Then
            Launcher.CreateProfile()
        End If

    End Sub

    Private Sub Brightness_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles Brightness.ValueChanged

        Dim brightness_res As Double = ((Brightness.Value - 1) / (30 - 1)) * 100

        If Brightness.Value = 30 Then
            Brightness_all = "Claro"
        ElseIf Brightness.Value = 1 Then
            Brightness_all = "Oscuro"
        Else
            Brightness_all = brightness_res.ToString("F0")
        End If

        lblBrightness.Content = "Brillo [" & Brightness_all & "]:"

    End Sub

    Private Sub MaxFPS_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles MaxFPS.ValueChanged

        If MaxFPS.Value = 200 Then
            Maxfps_all = "A tope"
        ElseIf MaxFPS.Value = 5 Then
            Maxfps_all = "VSync"
        Else
            Maxfps_all = MaxFPS.Value.ToString("F0")
        End If

        lblMaxFPS.Content = "FPS Máximos [" & Maxfps_all & "]:"

    End Sub

    Private Sub Profile_Name_ValueChanged(sender As Object, e As TextChangedEventArgs) Handles Profile_Name.TextChanged
        If Profile_Name.Text.Equals(String.Empty) Then
            btnSaveProfile.IsEnabled = False
        Else
            btnSaveProfile.IsEnabled = True
        End If
    End Sub

    Private Sub gridPLauncher_MLBD(sender As Object, e As MouseButtonEventArgs) Handles gridPLauncher.MouseLeftButtonDown
        ComboBoxW.Visibility = Windows.Visibility.Visible
        ComboBoxH.Visibility = Windows.Visibility.Visible
        txtBoxH.Visibility = Windows.Visibility.Hidden
        txtBoxW.Visibility = Windows.Visibility.Hidden
        brdWidth.Visibility = Windows.Visibility.Hidden
        brdHeight.Visibility = Windows.Visibility.Hidden
    End Sub

    Private Sub itemPreferncias_MLBD(sender As Object, e As MouseButtonEventArgs) Handles itemPreferencias.MouseLeftButtonDown
        ComboBoxW.SelectedIndex = 5
        ComboBoxH.SelectedIndex = 10

        Launcher.LoadAllLauncherSettings()

    End Sub

    Private Sub Button_Click_6(sender As Object, e As RoutedEventArgs)
        Launcher.SaveLauncherSettings()
    End Sub

    Private Sub login_Click_1(sender As Object, e As RoutedEventArgs) Handles login.Click
        Launcher.Log_in()
    End Sub

    Private Sub logout_Click(sender As Object, e As RoutedEventArgs) Handles logout.Click
        Launcher.Log_out()
    End Sub

    Private Sub tabSettings_Changed(sender As Object, e As SelectionChangedEventArgs) Handles tabCSettings.SelectionChanged

        Dim timer As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer() With {
            .Interval = TimeSpan.FromSeconds(0.01)
        }

        timer.Start()

        AddHandler timer.Tick, Sub()

                                   timer.Stop()

                                   If newFeatures.IsSelected Then

                                       If wbWindow IsNot Nothing Then
                                           wbWindow.Visibility = Windows.Visibility.Visible
                                       End If

                                   Else

                                       If wbWindow IsNot Nothing Then
                                           wbWindow.Visibility = Windows.Visibility.Hidden
                                       End If

                                   End If

                                   If itemOptions.IsSelected Then

                                       TSLoaded()

                                   End If

                               End Sub

        'If itemProEditor.IsSelected Then

        '    DataGridUtils.DeleteEmptyRows(dtProfiles)

        'End If

    End Sub

    Private Sub ComoResH_IChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBoxH.SelectionChanged
        If txtBoxH.Text IsNot "Altura" And txtBoxH.Text.Length > 0 Then
            txtBoxH.Text = ""
        End If
    End Sub

    Private Sub ComoResW_IChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBoxW.SelectionChanged
        If txtBoxW.Text IsNot "Anchura" And txtBoxW.Text.Length > 0 Then
            txtBoxW.Text = ""
        End If
    End Sub

    Private Sub dtControls_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles dtControls.SelectedCellsChanged
        'For Each cellInfo As DataGridCellInfo In dtControls.SelectedCells
        ' this changes the cell's content not the data item behind it

        '    Dim timer As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer() With {
        '    .Interval = TimeSpan.FromSeconds(0.01D)
        '}

        Dim totalRows As Integer = dtControls.Items.Count / dtControls.Columns.Count

        Dim CellsContents As List(Of String) = New List(Of String)

        Dim desiredColumnIndex As Integer = dtControls.Columns.Single(Function(c) c.Header = "Tecla/Botón").DisplayIndex

        For i As Integer = 0 To totalRows
            CellsContents.Add(DataGridHelper.GetCell(dtControls, i, desiredColumnIndex).Content.ToString())
        Next

        Dim cellInfo As DataGridCellInfo = dtControls.CurrentCell
        If cellInfo.Column.Header Is "Tecla/Botón" Then
            Dim gridCell As DataGridCell = DataGridHelper.TryToFindGridCell(dtControls, cellInfo)
            If gridCell IsNot Nothing Then
                If gridCell.Content IsNot "Press a key" And Not CellsContents.Any(Function(c) c = "Press a key") Then 'Check too if there is any cell that is going to be edited
                    gridCell.Content = "Press a key"
                End If
            End If
        End If
        'Next
    End Sub

    Private Sub OnKeyDownHandler(ByVal sender As Object, ByVal e As KeyEventArgs) Handles Me.KeyDown
        'KeyBoardHook = e.Key.ToString()

        If dtControls.IsFocused Then

            Dim cManager As ControlsManager = New ControlsManager

            Dim totalRows As Integer = dtControls.Items.Count / dtControls.Columns.Count

            Dim KeyIsTyped As Boolean = False

            Dim desiredColumnIndex As Integer = dtControls.Columns.Single(Function(c) c.Header = "Tecla/Botón").DisplayIndex

            Dim actionColumn As Integer = dtControls.Columns.Single(Function(c) c.Header = "Acción").DisplayIndex

            For i As Integer = 0 To totalRows
                If DataGridHelper.GetCell(dtControls, i, desiredColumnIndex).Content Is e.Key.ToString() Then
                    KeyIsTyped = True
                End If
            Next

            'Now, we can only edit that cell searching for cells that have inside them the keyword "Press a key"

            Dim cellInfo As DataGridCellInfo = dtControls.CurrentCell
            If cellInfo.Column.Header Is "Tecla/Botón" Then
                Dim gridCell As DataGridCell = DataGridHelper.TryToFindGridCell(dtControls, cellInfo)
                If gridCell IsNot Nothing Then
                    If gridCell.Content Is "Press a key" And Not KeyIsTyped Then 'Check if some cell has the same key...
                        gridCell.Content = e.Key.ToString()
                        cManager(Launcher.controls.IndexOf(DataGridHelper.GetCell(dtControls, Math.Floor(dtControls.SelectedIndex / dtControls.Columns.Count), actionColumn).Content)).Key = e.Key.ToString()
                    End If
                End If
            End If

        End If

    End Sub

    'Private Sub tbsettings_changed(sender As Object, e As SelectionChangedEventArgs) Handles tabSettings.SelectionChanged

    '    If itemControls.IsSelected And dtControls.Items.Count > 2 Then

    '        DataGridUtils.DeleteEmptyRows(dtControls)

    '    End If

    'End Sub

    Private Sub LoginButtonDisable()
        If Not String.IsNullOrEmpty(user.Text) And Not String.IsNullOrEmpty(pass.Password) Then
            login.IsEnabled = True
        Else
            login.IsEnabled = False
        End If
    End Sub

    Private Sub QualityLvl_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles QualityLvl.ValueChanged

        Dim qlvl As Integer = Convert.ToInt32(Math.Round(QualityLvl.Value))
        Dim finalStr As String = ""

        If qlvl = 5 Then
            finalStr = "Fantásticos"
        ElseIf qlvl = 4 Then
            finalStr = "Muy buenos"
        ElseIf qlvl = 3 Then
            finalStr = "Buenos"
        ElseIf qlvl = 2 Then
            finalStr = "Intermedios"
        ElseIf qlvl = 1 Then
            finalStr = "Malos"
        ElseIf qlvl = 0 Then
            finalStr = "Muy malos"
        Else
            finalStr = "*null*"
        End If

        lblQLvl.Content = "Nivel de calidad [" & finalStr & "]:"

    End Sub

#End Region

End Class

#End Region