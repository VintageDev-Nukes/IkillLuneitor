   M S B A M L   `   `   `  ���� O  Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem   GameLauncher WPF!   GameLauncher_WPF.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     9System&clr-namespace:System;assembly=mscorlib     I EWpfAnimatedGif, Version=1.4.4.0, Culture=neutral, PublicKeyToken=null,gif"http://wpfanimatedgif.codeplex.com    � Title$  
MainWindow��$	��823��$	��627�� � AllowsTransparency . 
������  � 
ResizeMode$ NoResize=� � WindowStyle$
 None=� � 	Resources  �� 
TargetType �� Property �� Value �� 
TargetType	 ��Name
 �� CornerRadius ��Name��%�   '�       �� " �.���  � ��$	 100����  � ��$	 100����  � ��$ 	10,0,10,0q���  � �� �� " �.R��� $	 
PART_Track��#��z��#��z��#��z��#��z��$
 5��._��� $ PART_Indicator��#��z��$
��Left=�#��z��.��� ����  �� Opacity$ 0Q� �� Color$ Black��.+��� $	��823��$	��627��$
��Left=�$	��Top=�����  �� 	UriSource$? 9pack://siteoforigin:,,,/Launcher/Resources/imgs/fondo.png?��� $	��110��$��20��$��613,497,0,0q�$
��Left=�$	��Top=�
�������
���   �$��1q�$
 5���� $	��110��$��20��$��613,472,0,0q�$
��Left=�$	��Top=�
�������
���   �$��1q�$
 5���� $	��110��$��20��$��613,472,0,0q�$
��Left=�$	��Top=����	#00636363���	#00000000$��1q�$��10(� �� TextWrapping$
 Wrap=��� $��82��$��20��$��526,472,0,0q�$
��Left=�$	��Top=�$��Center=�$��Usuario:��$��5,0,5,5q��� $��82��$��20��$��526,497,0,0q�$
��Left=�$	��Top=�$��Center=�$��Contraseña:��$��5,0,5,5q�G� $	��110��$��20��$��613,497,0,0q�$
��Left=�$	��Top=�
������ ���	#00000000�� -   $	��177��$��25��$��537,522,0,0q�$
��Left=�$	��Top=�$��Iniciar sesión��{�.  ����  �� 
StartPoint$	 0,0&� �� EndPoint$	 0,1&�.���  � Offset$ 0Q� � Color$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� -    ��Name$ cmbProfiles��$	��185��$��20��$��110,472,0,0q�$
��Left=�$	��Top=�$��Center=�$	��4,0q����� $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� $��40��$��20��$��
65,472,0,0q�$
��Left=�$	��Top=�$��Center=�$��Perfil:��$��0,-3,5,5q��� -   $	��100��$��25��$��
80,497,0,0q�$
��Left=�$	��Top=�$��Crear perfil������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� -   $	��100��$��25��$��188,497,0,0q�$
��Left=�$	��Top=�$��Editar perfil������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� -    ��Name$ tabCSettings��$	��400��$��39,54,51,173q����	#00F9F9F9���	#00DCDCDC  ��Name �� Triggers P� Property P� Value �� 
TargetName��%V  '��       �� " ��.���  � �� �� " ��.R��� $	 Border��$	��2,0q�
�������$��1,1,1,0q�$
 4,4,0,0��._��� $ ContentSite��$
��10,2q�$��Center=�$��Center=�$��Header�� P�  �@��
IsSelected . . ��� $ Border�� � ��
 ��·�P�  �@��
IsSelected .  . ��� $ Border�� � ��
 �����.@��� $��Nuevas características��.���� -    ��Name$ 
LayoutRoot��
�������.+�(  "System.Windows.Controls.WebBrowser  -     Name$ browser�� �� Left$ 10��
 �� Top$ 10��$	��703��$	��347��   Source$ http://google.es?��� $��Editor de perfiles��.��&   System.Windows.Controls.DataGrid  -      Name$  
dtProfiles��$	��725��$	��360��$��0,10,-2,-23q�$
��Left=�$	��Top=�$!   HorizontalScrollBarVisibility$! Disabled=��� $��Opciones de juego��.��� $	��719��$	��373��$��0,0,0,0q�$
��Left=�$	��Top=�.+��� $��122,10,457,343q�
�������
���   �$��1q�$
 5���� $��16,10,602,343q�$��Nombre del perfil:��$��0q��� -	   " ��Name$" Profile_Name��$��122,10,457,343q�$��0q�$
 Wrap=���Z� ��Z� �� -
   $	��699��$��10,35q�$��Center=�.@��� $��50��$��0q�$	��Top=�$��Stretch=�$��Stretch=�$	��6,1q����� .R��� # �� Orientation$# 
Horizontal=�.+��� $��25��$��25��$C��=pack://siteoforigin:,,,/Launcher/Resources/icons/Keyboard.png��$
��Fill=��� $��5,5,0,0q�$��	Controles��.��� .+�  $	��669��$��
10,10,0,10q�$
��Left=�$! Disabled=�$   Columns$ 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn%   Width5  /System.Windows.Controls.DataGridLengthConverter$	% 335 &   Header$& Acción��'   
IsReadOnly' . 1  +System.Windows.Controls.DataGridBoundColumn(   Binding( Z� )   ClipboardContentBinding) Z�   $	% 335 $& Tecla/Botón��' . ( Z� ) Z� �� $��50��$��0q�$
��Left=�$	��Top=�$��Right=����� .R��� $# 
Horizontal=�.+��� $��25��$��25��$@��:pack://siteoforigin:,,,/Launcher/Resources/icons/Mouse.png��$
��Fill=��� $��5,5,0,0q�$��Ratón��.��� $	��688��.+��� -   * ��Name$* lblSens��$��25��$��10,10,464,206q�$��Sensibilidad:��$��5q��� -   + ��Name$+ Sensibility��$��25��$��10,0,354,176q�$��Bottom=�, � Maximum$, 30Q�- � Minimum$	- 0.1Q��� -   $+ Sensibility_Copy��$��25��$��353,40,10,176q�$, 30Q�$	- 0.1Q�$y�Hidden=��� -   . ��Name$. InvertMouse��$��89��$��15��$��	10,70,0,0q�$
��Left=�$	��Top=�$��Invertir ratón���� $��50��$��0q�$
��Left=�$	��Top=�$��Right=����� .R��� $# 
Horizontal=�.+��� $��25��$��25��$@��:pack://siteoforigin:,,,/Launcher/Resources/icons/Audio.png��$
��Fill=��� $��5,5,0,0q�$��Audio��.��� �� $��50��$��0q�$
��Left=�$	��Top=�$��Right=����� .R��� $# 
Horizontal=�.+��� $��25��$��25��$@��:pack://siteoforigin:,,,/Launcher/Resources/icons/Video.png��$
��Fill=��� $��5,5,0,0q�$��Video��.��� $	��690��.+��� -   $* lblFOV��$��25��$��10,10,464,206q�$��FOV (Field Of View):��$��5q��� -   $	+ FOV��$��25��$��10,0,354,176q�$��Bottom=�$	, 120Q�$- 60Q��� -   $* lblBrightness��$��25��$��353,10,121,206q�$��Brillo:��$��5q��� -   $* lblRenderDistance��$��25��$��10,70,464,146q�$��Distancia de renderizado:��$��5q��� -   $* 	lblMaxFPS��$��25��$��353,70,121,146q�$��FPS Máximos:��$��5q��� -   $+ 
Brightness��$	��327��$��25��$��
353,40,0,0q�$
��Left=�$	��Top=�$, 30Q�$- 1Q��� -   $+ MaxFPS��$	��327��$��25��$��353,100,0,0q�$
��Left=�$	��Top=�$	, 200Q�$- 5Q��� -   $+ RenderDistance��$	��326��$��25��$��
10,100,0,0q�$
��Left=�$	��Top=�$
, 8000Q�$	- 100Q��� $��50��$
��Left=�$	��Top=����� .R��� $# 
Horizontal=�.+��� $��25��$��25��$C��=pack://siteoforigin:,,,/Launcher/Resources/icons/Settings.png��$
��Fill=��� $��5,5,0,0q�$��Otros��.��� �� -   / ��Name$/ btnSaveProfile��$	��148��$��25��$��285,343,284,0q�$��Center=�$	��Top=�$��Guardar perfil������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� -   0 ��Name$0 itemPreferencias��$��21��$
��Left=�$	��Top=�$��Prefencias del Launcher��.��� -   1 �Name$1 gridPLauncher��
�������.+��� $	��703��$��68��$��	10,10,0,0q�$
��Left=�$	��Top=�$��
Apariencia��.��� .+��� $��46��$��25��$��	10,10,0,0q�$
��Left=�$	��Top=�$��Tema:���� $��61,10,463,11q����� $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� $	��703��$��68��$��	10,83,0,0q�$
��Left=�$	��Top=�$��Lenguaje��.��� .+��� $��72��$��25��$��	10,10,0,0q�$
��Left=�$	��Top=�$��	Lenguaje:���� $��87,10,437,11q����� $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� $	��703��$	��108��$��
10,151,0,0q�$
��Left=�$	��Top=�$��Opciones de juego��.��� .+��� -   $	 brdWidth��$��74��$��25��$��	87,10,0,0q�$
��Left=�$	��Top=�
�������
���   �$��1q�$
 5��$y�Hidden=��� -   $	 	brdHeight��$��74��$��25��$��
176,10,0,0q�$
��Left=�$	��Top=�
�������
���   �$��1q�$
 5��$y�Hidden=��� -   $" txtBoxW��$��74��$��25��$��	87,10,0,0q�$
��Left=�$	��Top=�
���   �$��Anchura��$
 Wrap=�$y�Hidden=���Z� ��Z� �� -   $" txtBoxH��$��74��$��25��$��
176,10,0,0q�$
��Left=�$	��Top=�$��Altura��$
 Wrap=�$y�Hidden=���Z� ��Z� �� $��0,10,609,51q�$��Resolución:���� -   $ 	ComboBoxH��$��176,10,441,51q�$��Center=�2 � ToolTip$=2 7Haga doble clic para introducir un valor predeterminado������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White��.@���2048����1600����1536����1200����1080����1050����1024����960����900����854����768����720����600����576����480����240����200���� $��161,11,515,50q�$��Center=�$��Center=�$��x��$��0q��� -   $ 	ComboBoxW��$��87,10,530,51q�$��Center=�$=2 7Haga doble clic para introducir un valor predeterminado������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White��.@���2560����2048����1920����1600����1400����1366����1280����1152����1024����854����800����768����640����320���� -    $. chkFullScreen��$	��113��$��16��$��	10,40,0,0q�$
��Left=�$	��Top=�$��Pantalla completa���� -!   $	��175��$��25��$��
10,332,0,0q�$��Center=�$	��Top=�$��Guardar preferencias������ $	 0,0&�$	 0,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� -"   $��24��$��24��$��
776,10,0,0q�$
��Left=�$	��Top=�$��Center=�
���   �
���   �$��0,0,1,1q�$��x��$��20(�$
��Bold"����	#FF00DCFF$��2,-7,1,1q��� -#   $/ btnPlay��$	��148��$��55��$��332,477,0,0q�$
��Left=�$	��Top=�$��JUGAR��$��25(�$��	ExtraBold"�
���   ����� 3 � MappingMode$3 RelativeToBoundingBox=�$ 0.5,0&�$ 0.5,1&�.��� $ 0Q�$ White��� $	 0.5Q�$
 Gray��� $ 1Q�$ White���� $��10,552,23,42q�$# 
Horizontal=�.+��� -$   4 ��Name$4 animProgress��$��25��$��25��$��0q�$y�Hidden=�"  WpfAnimatedGif.ImageBehavior5   AnimatedSource$E5 ?pack://siteoforigin:,,,/Launcher/Resources/imgs/progressbar.gif��6   RepeatBehavior$6 Forever���� -%   $4 imgCompleted��$��25��$��25��$��	-25,0,0,0q�$C��=pack://siteoforigin:,,,/Launcher/Resources/imgs/tick-icon.png��$y�Hidden=��� -&   $* lblProgress��$	��759��$��20��$��5,10,0,0q�$
��Left=�$	��Top=�$��	Progreso:��
�������$	��5,0q��� $# 
Horizontal=�.+��� ���� -'   img��7 j� Stretch$7 UniformToFill=�8 j� TileMode$
8 Tile=�9 j� Viewport$9 	0,0,10,10�: j� ViewportUnits$: Absolute=�� -(   ; �Name$; ProgressBar1��$	��790��$��10��$��
10,584,0,0q�$
��Left=�$	��Top=����	#FF292929���	#FF838383< J� IsHitTestVisible< . 
= J� Uid$= ProgressBar1��> � Value$	> 100Q����� ? �� ElementName$	? img���� $	��249��$��39��$��	10,10,0,0q�$
��Left=�$	��Top=�$>��8pack://siteoforigin:,,,/Launcher/Resources/imgs/Logo.png��