Imports System.IO
Imports System.Net
Imports System.Linq

Public Class SplashText

    Dim wc As WebClient
    Dim FinalGamePath As String
    Dim currentPath As String

    ReadOnly LauncherKey As String = "1"

    Private Shadows Sub Load() Handles MyBase.Loaded

        currentPath = System.AppDomain.CurrentDomain.BaseDirectory

        lblEstado.Content = "Definiendo variables..."

        Dim width As Double = System.Windows.SystemParameters.PrimaryScreenWidth
        Dim height As Double = System.Windows.SystemParameters.PrimaryScreenHeight
        SplashText.Left = width / 2 - SplashText.Width / 2
        SplashText.Top = height / 2 - SplashText.Height / 2

        Logo.Source = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/Logo.png"))
        pbEstadoImg.ImageSource = New BitmapImage(New Uri("pack://application:,,,/Resources/imgs/pbc.png"))

        Dim AppData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        'Pienso que al final para hacer dos comprobaciones, si lo del .net shrink va, junto el splash text con el launcher

        lblEstado.Content = "Comprobando si hay actualizaciones disponibles..."

        Dim FixedURLCheckerLauncher = "http://dl.dropboxusercontent.com/s/n1aemm3wwdeboaa/Dynawars%20Launcher%20Key%20Updater.txt?dl=1"

        'Comprobar si hay actualizaciones para el Launcher...
        'Hacer un archivo SFX que se ejecute en segundo plano
        Updater.CheckAndDownload(FixedURLCheckerLauncher, AppData & "\.dynawars\temp\LauncherUpt.exe", LauncherKey)

        'Generate a key for the SFX file
        File.AppendAllText(AppData & "\.dynawars\temp\SFXExtraction.txt", Directory.GetCurrentDirectory())

        lblEstado.Content = "Haciendo unas pequeñas comprobaciones..."

        'Create Launcher and game folders

        IOUtils.CreateDirectoryFromRoot(AppData & "\.dynawars\Launcher")

        IOUtils.CreateDirectoryFromRoot(AppData & "\.dynawars\temp")

        'Checks if there is an old configuration (if not ask to the user where to install the game)

        Dim configexists As Boolean = File.Exists(AppData & "\.dynawars\app.conf") 'Checks if a old config exists
        Dim foldersetskipped As Boolean = True 'This will be the bool that set if something of those options was changed
        Dim GamePath As String = "" 'This will contain the path of the game

        If Not configexists Then 'If not configuration exists before...

            File.Create(AppData & "\.dynawars\app.conf") 'Create it

            'And set where to install
            FinalGamePath = IOUtils.NewGamePath("Seleccione una ruta para instalar el juego")

            IOUtils.CreateDirectoryFromRoot(FinalGamePath)
            IOUtils.ValidatePath(FinalGamePath)

            INIFileManager.Key.Set("GamePath", FinalGamePath)

            foldersetskipped = False 'Don't forget to say to the script that we changed something

        Else 'If a config exists before...

            INIFileManager.FilePath = AppData & "\.dynawars\app.conf"

            Dim keyexists As Boolean = INIFileManager.Key.Exist("GamePath") 'Checks if an old configuration contain the gamepath or if it is disfigured completely

            If keyexists Then 'If key exists...

                GamePath = INIFileManager.Key.Get("GamePath").ToString()

                'Checks if the directory that the key contains still exists...
                If Not Directory.Exists(GamePath) Then 'If not...

                    'Set a new path for the game
                    FinalGamePath = IOUtils.NewGamePath("La carpeta del juego que había en la configuración no existe, por favor, de una ruta nueva de instalación")

                    IOUtils.CreateDirectoryFromRoot(FinalGamePath)
                    IOUtils.ValidatePath(FinalGamePath)

                    INIFileManager.Key.Set("GamePath", FinalGamePath)

                    foldersetskipped = False 'And say to the that we change something...

                Else 'If the directory still exists...

                    FinalGamePath = GamePath 'Set the variable

                End If


            Else 'If the key doesn't exists set a new path for the game

                FinalGamePath = IOUtils.NewGamePath("La configuración del juego no se encuentra, por favor, elija una nueva ruta para instalarlo")

                IOUtils.CreateDirectoryFromRoot(FinalGamePath)
                IOUtils.ValidatePath(FinalGamePath)

                INIFileManager.Key.Set("GamePath", FinalGamePath)

                foldersetskipped = False 'And say that to the script that we changed something

            End If

        End If

        'Set Launcher path for the game

        INIFileManager.Key.Set("LauncherPath", System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName)

        'Check all the files from the game

        If Not Directory.Exists(FinalGamePath & "\Game_Data") And foldersetskipped Then 'If the game is corrupt and we didn't changed anything before in the config

            FinalGamePath = IOUtils.NewGamePath("El juego está corrupto o no se instaló correctamente, por favor, elija una nueva ruta de instalación")

            IOUtils.CreateDirectoryFromRoot(FinalGamePath)
            IOUtils.ValidatePath(FinalGamePath)

            INIFileManager.Key.Set("GamePath", FinalGamePath)

        End If

        'If everything is correct then set the variable

        Paths.GamePath = INIFileManager.Key.Get("GamePath")

        'Continue showing the Launcher

        lblEstado.Content = "Mostrando el Launcher..."

        'FormStyle.FormFade("out", Window.GetWindow(Me))

        Dim mw = New MainWindow()

        mw.Show()

        Me.Close()

    End Sub

    Private Sub pbEstado_Loaded(sender As Object, e As RoutedEventArgs) Handles pbEstado.Loaded
        Dim p = DirectCast(sender, ProgressBar)
        p.ApplyTemplate()

        DirectCast(p.Template.FindName("Animation", p), Panel).Background = Brushes.Transparent
    End Sub

End Class
