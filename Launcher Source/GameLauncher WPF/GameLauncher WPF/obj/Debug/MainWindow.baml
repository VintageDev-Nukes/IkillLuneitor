   M S B A M L   `   `   `  ����O  Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem  5   	    GameLauncher WPF!   GameLauncher_WPF.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   9System&clr-namespace:System;assembly=mscorlib     5   	   I EWpfAnimatedGif, Version=1.4.4.0, Culture=neutral, PublicKeyToken=null,gif"http://wpfanimatedgif.codeplex.com  5   	     � Title$  
MainWindow��5   	   $	��823��5   	   $	��627��5   	    � AllowsTransparency . 
������ ���	#00000000 � 
ResizeMode$ NoResize=�5   	    � WindowStyle$
 None=�5   	    � 	Resources 5       �� 
TargetType �� Property �� Value �� 
TargetType	 ��Name
 �� CornerRadius ��Name��%�  '�       �� 5   
   " �6   .�5      ��  � ��$	 100��6&   �� 5       � ��$	 100��6'   �� 5       � ��$ 	10,0,10,0q�6'   �� 5       � �� 5      �� 5      " �6&   .R�5      �� $	 
PART_Track��6!   #��z��5   !   #��z��5   !   #��z��5   !   #��z��5   !   $
 5��5   !   ._�5      �� $ PART_Indicator��6(   #��z��5    (   $
��Left=�5!   (   #��z��5"   (   5#      5%      5'      5(      5)      .��5+      � ��5,   
   �� 5-       �� Opacity$ 0Q�6    �� Color$ Black��6*   .+�5/   
   �� -   5         ��Name$ Fondo��5/      $	��823��50      $	��627��51      $
��Left=�52      $	��Top=�53      �� 54   
   -   5         ��Name$ cmbProfiles��54      $	��185��55      $��20��56      $��110,472,0,0q�57      $
��Left=�58      $	��Top=�59      $��Center=�5:      $	��4,0q�5;      ��5<      �� 5=       �� 
StartPoint$	 0,0&�6&    �� EndPoint$	 0,1&�67   .��5>      �  � Offset$ 0Q�6#    � Color$ White��6.   � 5?      $	 0.5Q�6#   $
 Gray��60   � 5@      $ 1Q�6#   $ White��6.   5A      5C      �� 5D   
   $��40��6   $��20��5E      $��
65,472,0,0q�5F      $
��Left=�5G      $	��Top=�5H      $��Center=�5I      $��Perfil:��5J      $��0,-3,5,5q�5K      �� 5L   
   -   5        $	��100��5L      $��25��5M      $��
80,497,0,0q�5N      $
��Left=�5O      $	��Top=�5P      $��Crear perfil��5R      ��5S      �� 5T      $	 0,0&�6&   $	 0,1&�67   .��5U      � $ 0Q�6#   $ White��6.   � 5V      $	 0.5Q�6#   $
 Gray��60   � 5W      $ 1Q�6#   $ White��6.   5X      5Z      �� 5[   
   -   5        $	��100��5[      $��25��5\      $��188,497,0,0q�5]      $
��Left=�5^      $	��Top=�5_      $��Editar perfil��5a      ��5b      �� 5c      $	 0,0&�6&   $	 0,1&�67   .��5d      � $ 0Q�6#   $ White��6.   � 5e      $	 0.5Q�6#   $
 Gray��60   � 5f      $ 1Q�6#   $ White��6.   5g      5i      �� 5j   
   -   5         ��Name$ tabCSettings��5j      $	��400��5k      $��39,54,51,173q�5l      ���	#00F9F9F9���	#00DCDCDC 5o       ��Name �� Triggers P� Property P� Value �� 
TargetName��%V  '��       �� 5p      " ��6   .�5q      ��  � �� 5r      �� 5s      " ��6.   .R�5t   "   �� $	 Border��6)   $	��2,0q�5u   )   
�������$��1,1,1,0q�5w   )   $
 4,4,0,0��5x   )   ._�5y   &   �� $ ContentSite��67   $
��10,2q�5z   7   $��Center=�5{   7   $��Center=�5|   7   $��Header��5}   7   5~   #    5   "   P� 5�   &    �@��
IsSelected . . �5�   *   �� $ Border��61    � ��
 ��·�6[   5�   '   P� 5�   &    �@��
IsSelected .  . �5�   *   �� $ Border��61    � ��
 �����6[   5�   '   5�      5�      5�      5�      .@�5�      �� -   5         ��Name$ newFeatures��5�      $��Nuevas características��6+   .��5�      �� -   5         ��Name$ 
LayoutRoot��5�      
������ .+�5�      �� -   5        $ ��_webBrowserPlacementTarget��5�       �� Left$ 10��5�      
 �� Top$ 10��5�      $	��703��5�      $	��347��5�      
����  �$��2q�5�      $	��Top=�5�      5�      5�      �� 5�      -	   5        $ itemProEditor��5�      $��Editor de perfiles��6-   .��5�      &   System.Windows.Controls.DataGrid  -
   5          Name$ 
dtProfiles��5�      $	��725��5�      $	��360��5�      $��0,10,-2,-23q�5�      $
��Left=�5�      $	��Top=�5�      $   HorizontalScrollBarVisibility$ Disabled=�5�      5�      5�      �� 5�      -   5        $ itemOptions��5�      $��Opciones de juego��6+   .��5�      � -   5          �Name$  gameOptions��5�      $	��719��5�      $	��373��5�      $��0,0,0,0q�5�      $
��Left=�5�      $	��Top=�5�      .+�5�      �� $��122,10,457,343q�6   
�������
���   �$��1q�5�      $
 5��5�      �� 5�      $��16,10,602,343q�6   $��Nombre del perfil:��5�      $��0q�5�      �� 5�      -   5        ! ��Name$! Profile_Name��5�      $��122,10,457,343q�5�      $��0q�5�      " �� TextWrapping$
" Wrap=�5�      ��5�      Z� ��5�      Z� 5�      �� 5�      -   5        $ tabSettings��5�   !   $	��699��5�   !   $��10,35q�5�   !   $��Center=�5�   !   .@�5�      �� -   5        $ itemControls��5�   "   $��50��5�   "   $��0q�5�   "   $	��Top=�5�   "   $��Stretch=�5�   "   $��Stretch=�5�   "   $	��6,1q�5�   "   ��5�      �� 5�   "   .R�5�   &   �� # �� Orientation$# 
Horizontal=�61   .+�5�   *   �� $ iconKB��60   $��25��5�   0   $��25��5�   0   $
��Fill=�5�   0   �� 5�   *   $��5,5,0,0q�64   $��	Controles��6E   5�   '   5�   #   .��5�      � .+�5�   "     -   5        $ 
dtControls��5�   +   $	��669��5�   +   $��
10,10,0,10q�5�   +   $
��Left=�5�   +   $ Disabled=�5�   +   5�   #   5�      5�      �� 5�      -   5        $ 	itemMouse��5�   "   $��50��5�   "   $��0q�5�   "   $
��Left=�5�   "   $	��Top=�5�   "   $��Right=�5�   "   ��5�      �� 5�   "   .R�5�   &   �� $# 
Horizontal=�61   .+�5�   *   �� $ 	iconMouse��60   $��25��5�   0   $��25��5�   0   $
��Fill=�5�   0   �� 5�   *   $��5,5,0,0q�64   $��Ratón��6E   5�   '   5�   #   .��5�      � $	��688��6#   .+�5�   "   �� -   5        $ ��Name$$ lblSens��5�   (   $��25��5   (   $��10,10,464,206q�5  (   $��Sensibilidad:��5  (   $��5q�5  (   �� 5  "   -   5        % ��Name$% Sensibility��5  )   $��25��5  )   $��10,0,354,176q�5  )   $��Bottom=�5  )   & � Maximum$& 30Q�5  )   ' � Minimum$	' 0.1Q�5	  )   5  )   �� 5  "   -   5        $% Sensibility_Copy��5  )   $��25��5  )   $��353,40,10,176q�5  )   $& 30Q�5  )   $	' 0.1Q�5  )   $y�Hidden=�5  )   �� 5  "   -   5        ( ��Name$( InvertMouse��5  +   $��89��5  +   $��15��5  +   $��	10,70,0,0q�5  +   $
��Left=�5  +   $	��Top=�5  +   $��Invertir ratón��5  +   5  +   5     5     �� 5     -   5        $ 	itemAudio��5  "   $��50��5  "   $��0q�5  "   $
��Left=�5   "   $	��Top=�5!  "   $��Right=�5"  "   ��5#     �� 5$  "   .R�5%  &   �� $# 
Horizontal=�61   .+�5&  *   �� $ 	iconAudio��60   $��25��5'  0   $��25��5(  0   $
��Fill=�5)  0   �� 5*  *   $��5,5,0,0q�64   $��Audio��6E   5+  '   5,  #   .��5.     � 5/     �� 50     -   5        $ 	itemVideo��50  "   $��50��51  "   $��0q�52  "   $
��Left=�53  "   $	��Top=�54  "   $��Right=�55  "   ��56     �� 57  "   .R�58  &   �� $# 
Horizontal=�61   .+�59  *   �� $ 	iconVideo��60   $��25��5:  0   $��25��5;  0   $
��Fill=�5<  0   �� 5=  *   $��5,5,0,0q�64   $��Video��6E   5>  '   5?  #   .��5A     � $	��690��6#   .+�5K  "   �� $��10,130,354,45q�6+   $��Resolución��6B   .��5L  &   � $��0,0,0,-1q�6+   .+�5M  *   �� -   5        $	 brdWidth��5M  1   $��60��5N  1   $��25��5O  1   $��	10,12,0,0q�5P  1   $
��Left=�5Q  1   $	��Top=�5R  1   
�������
���   �$��1q�5U  1   $
 5��5V  1   $y�Hidden=�5W  1   �� 5X  *   -   5        $	 	brdHeight��5X  1   $��60��5Y  1   $��25��5Z  1   $��	85,12,0,0q�5[  1   $
��Left=�5\  1   $	��Top=�5]  1   
�������
���   �$��1q�5`  1   $
 5��5a  1   $y�Hidden=�5b  1   �� 5c  *   -   5        $! txtBoxW��5c  2   $��60��5d  2   $��25��5e  2   $��	10,12,0,0q�5f  2   $
��Left=�5g  2   $	��Top=�5h  2   
���   �$��Anchura��5l  2   $
" Wrap=�5m  2   $y�Hidden=�5n  2   ��5i  2   Z� ��5j  2   Z� 5n  2   �� 5o  *   -   5        $! txtBoxH��5o  2   $��60��5p  2   $��25��5q  2   $��	85,12,0,0q�5r  2   $
��Left=�5s  2   $	��Top=�5t  2   $��Altura��5w  2   $
" Wrap=�5x  2   $y�Hidden=�5y  2   ��5u  2   Z� ��5v  2   Z� 5y  2   �� 5z  *   -   5        $ 	ComboBoxH��5z  3   $��60��5{  3   $��25��5|  3   $��85,10,169,10q�5}  3   $��Center=�5~  3   ) � ToolTip$=) 7Haga doble clic para introducir un valor predeterminado��5�  3   ��5�  .   �� 5�  2   $	 0,0&�6F   $	 0,1&�6W   .��5�  6   � $ 0Q�6C   $ White��6N   � 5�  6   $	 0.5Q�6C   $
 Gray��6P   � 5�  6   $ 1Q�6C   $ White��6N   5�  3   .@�5�  .   ��2048��6A   ��5�  .   1600��6A   ��5�  .   1536��6A   ��5�  .   1200��6A   ��5�  .   1080��6A   ��5�  .   1050��6A   ��5�  .   1024��6A   ��5�  .   960��6@   ��5�  .   900��6@   ��5�  .   854��6@   ��5�  .   768��6@   ��5�  .   720��6@   ��5�  .   600��6@   ��5�  .   576��6@   ��5�  .   480��6@   ��5�  .   240��6@   ��5�  .   200��6@   5�  +   �� 5�  *   $��15��60   $��70,10,229,10q�5�  0   $��Center=�5�  0   $��Center=�5�  0   $��x��5�  0   $��0q�5�  0   �� 5�  *   -   5        $ 	ComboBoxW��5�  3   $��60��5�  3   $��25��5�  3   $��10,10,244,10q�5�  3   $��Center=�5�  3   $=) 7Haga doble clic para introducir un valor predeterminado��5�  3   ��5�  .   �� 5�  2   $	 0,0&�6F   $	 0,1&�6W   .��5�  6   � $ 0Q�6C   $ White��6N   � 5�  6   $	 0.5Q�6C   $
 Gray��6P   � 5�  6   $ 1Q�6C   $ White��6N   5�  3   .@�5�  .   ��2560��6A   ��5�  .   2048��6A   ��5�  .   1920��6A   ��5�  .   1600��6A   ��5�  .   1400��6A   ��5�  .   1366��6A   ��5�  .   1280��6A   ��5�  .   1152��6A   ��5�  .   1024��6A   ��5�  .   854��6@   ��5�  .   800��6@   ��5�  .   768��6@   ��5�  .   640��6@   ��5�  .   320��6@   5�  +   5�  '   5�  #   �� 5�  "   -   5        $( chkFullScreen��5�  +   $	��113��5�  +   $��16��5�  +   $��186,162,0,0q�5�  +   $
��Left=�5�  +   $	��Top=�5�  +   $��Pantalla completa��5�  +   5�  +   �� 5�  "   -   5        $$ lblFOV��5�  (   $��25��5�  (   $��10,10,464,206q�5�  (   $��FOV (Field Of View):��5�  (   $��5q�5�  (   �� 5�  "   -    5        $	% FOV��5�  )   $��25��5�  )   $��10,0,354,176q�5�  )   $��Bottom=�5�  )   $	& 120Q�5�  )   $' 60Q�5�  )   5�  )   �� 5�  "   -!   5        $$ lblBrightness��5�  (   $��25��5�  (   $��353,10,121,206q�5�  (   $��Brillo:��5�  (   $��5q�5�  (   �� 5�  "   -"   5        $$ lblRenderDistance��5�  (   $��25��5�  (   $��10,70,464,146q�5�  (   $��Distancia de renderizado:��5�  (   $��5q�5�  (   �� 5�  "   -#   5        $$ 	lblMaxFPS��5�  (   $��25��5�  (   $��353,70,121,146q�5�  (   $��FPS Máximos:��5�  (   $��5q�5�  (   �� 5�  "   -$   5        $% 
Brightness��5�  )   $	��327��5�  )   $��25��5�  )   $��
353,40,0,0q�5�  )   $
��Left=�5�  )   $	��Top=�5�  )   $& 30Q�5�  )   $' 1Q�5�  )   5�  )   �� 5�  "   -%   5        $% MaxFPS��5�  )   $	��327��5�  )   $��25��5�  )   $��353,100,0,0q�5�  )   $
��Left=�5�  )   $	��Top=�5�  )   $	& 200Q�5�  )   $' 5Q�5�  )   5�  )   �� 5�  "   -&   5        $% RenderDistance��5�  )   $	��326��5�  )   $��25��5�  )   $��
10,100,0,0q�5�  )   $
��Left=�5   )   $	��Top=�5  )   $
& 8000Q�5  )   $	' 100Q�5  )   5  )   �� 5  "   -'   5        $$ lblQLvl��5  (   $	��327��69   $��25��5  (   $��353,130,0,0q�5  (   $
��Left=�5  (   $	��Top=�5	  (   $��Nivel de calidad:��5
  (   �� 5  "   -(   5        $% 
QualityLvl��5  )   $	��327��5  )   $��26��5  )   $��353,160,0,0q�5  )   $
��Left=�5  )   $	��Top=�5  )   $& 5Q�5  )   5     5     �� 5     -)   5        $ itemSettings��5  "   $��50��5  "   $
��Left=�5  "   $	��Top=�5  "   ��5     �� 5  "   .R�5  &   �� $# 
Horizontal=�61   .+�5  *   �� $ iconSettings��60   $��25��5  0   $��25��5  0   $
��Fill=�5  0   �� 5  *   $��5,5,0,0q�64   $��Otros��6E   5   '   5!  #   .��5#     � .+�5$  "   �� $��72��6(   $��30��5%  (   $��	10,10,0,0q�5&  (   $
��Left=�5'  (   $	��Top=�5(  (   $��Center=�5)  (   $��	Lenguaje:��5*  (   �� 5+  "   $��87,10,482,201q�6+   ��5,  &   �� 5-  *   $	 0,0&�6>   $	 0,1&�6O   .��5.  .   � $ 0Q�6;   $ White��6F   � 5/  .   $	 0.5Q�6;   $
 Gray��6H   � 50  .   $ 1Q�6;   $ White��6F   51  +   53  #   54     55     56     �� 57     -*   5        * ��Name$* btnSaveProfile��57     $	��148��58     $��25��59     $��285,343,284,0q�5:     $��Center=�5;     $	��Top=�5<     $��Guardar perfil��5>     ��5?     �� 5@     $	 0,0&�62   $	 0,1&�6C   .��5A  "   � $ 0Q�6/   $ White��6:   � 5B  "   $	 0.5Q�6/   $
 Gray��6<   � 5C  "   $ 1Q�6/   $ White��6:   5D     5F     5H     5I     �� 5J     -+   5        $ itemPreferencias��5J     $��21��5K     $
��Left=�5L     $	��Top=�5M     $��Preferencias del Launcher��5N     .��5O     � -,   5        $  gridPLauncher��5O     
�������.+�5P     �� $	��703��6   $��68��5Q     $��	10,10,0,0q�5R     $
��Left=�5S     $	��Top=�5T     $��
Apariencia��5U     .��5V     � .+�5W     �� $��46��6$   $��25��5X  $   $��	10,10,0,0q�5Y  $   $
��Left=�5Z  $   $	��Top=�5[  $   $��Tema:��5\  $   �� 5]     $��61,10,463,11q�6'   ��5^  "   �� 5_  &   $	 0,0&�6:   $	 0,1&�6K   .��5`  *   � $ 0Q�67   $ White��6B   � 5a  *   $	 0.5Q�67   $
 Gray��6D   � 5b  *   $ 1Q�67   $ White��6B   5c  '   5e     5f     5g     �� 5h     $	��703��6   $��68��5i     $��	10,83,0,0q�5j     $
��Left=�5k     $	��Top=�5l     $��Lenguaje��5m     .��5n     � .+�5o     �� $��72��6$   $��25��5p  $   $��	10,10,0,0q�5q  $   $
��Left=�5r  $   $	��Top=�5s  $   $��	Lenguaje:��5t  $   �� 5u     $��87,10,437,11q�6'   ��5v  "   �� 5w  &   $	 0,0&�6:   $	 0,1&�6K   .��5x  *   � $ 0Q�67   $ White��6B   � 5y  *   $	 0.5Q�67   $
 Gray��6D   � 5z  *   $ 1Q�67   $ White��6B   5{  '   5}     5~     5     �� 5�     --   5        $	��175��5�     $��25��5�     $��
10,332,0,0q�5      $��Center=�5     $	��Top=�5     $��Guardar preferencias��5     ��5     �� 5     $	 0,0&�62   $	 0,1&�6C   .��5  "   � $ 0Q�6/   $ White��6:   � 5  "   $	 0.5Q�6/   $
 Gray��6<   � 5	  "   $ 1Q�6/   $ White��6:   5
     5     5     5     5     �� 5  
   -.   5        $* closeApp��5     $��24��5     $��24��5     $��
776,10,0,0q�5     $
��Left=�5     $	��Top=�5     $��Center=�5     
���   �
���   �$��0,0,1,1q�5     $��x��5     $��20(�5     $
��Bold"�5     ���	#FF00DCFF$��2,-7,1,1q�5     �� 5   
   -/   5        $* btnPlay��5      $	��148��5!     $��55��5"     $��332,477,0,0q�5#     $
��Left=�5$     $	��Top=�5%     $��JUGAR��5'     $��25(�5(     $��	ExtraBold"�5)     
���   ���5+     �� 5,     + � MappingMode$+ RelativeToBoundingBox=�6&   $ 0.5,0&�6J   $ 0.5,1&�6]   .��5-     � $ 0Q�6#   $ White��6.   � 5.     $	 0.5Q�6#   $
 Gray��60   � 5/     $ 1Q�6#   $ White��6.   50     52     �� 5G  
   $��10,552,23,42q�6   $# 
Horizontal=�6+   .+�5I     �� -0   5        $ animProgress��5I     $��25��5J     $��25��5K     $��0q�5L     $y�Hidden=�5M     �� 5O     -1   5        $ imgCompleted��5O     $��25��5P     $��25��5Q     $��	-25,0,0,0q�5R     $y�Hidden=�5S     �� 5U     -2   5        $$ lblProgress��5U     $	��759��5V     $��20��5W     $��5,10,0,0q�5X     $
��Left=�5Y     $	��Top=�5Z     $��	Progreso:��5[     
�������$	��5,0q�5]     5^     �� 5a  
   $# 
Horizontal=�6   .+�5c     �� ��5d     �� 5e     -3   5        img��, j� Stretch$, UniformToFill=�5f  !   - j� TileMode$
- Tile=�5g  !   . j� Viewport$. 	0,0,10,10�5h  !   / j� ViewportUnits$/ Absolute=�5i  !   5k     � 5l     -4   5        0 �Name$0 ProgressBar1��5l     $	��790��5m     $��10��5n     $��
10,584,0,0q�5o     $
��Left=�5p     $	��Top=�5q     ���	#FF292929���	#FF8383831 J� IsHitTestVisible1 . 2 � IsIndeterminate2 . 
3 J� Uid$3 ProgressBar1��5w     4 � Value$	4 100Q�5x     ��5t     �� 5 �� ElementName$	5 img��5x     5y     �� 5z  
   -5   5        $
 Logo��5z     $	��249��5{     $��39��5|     $��	10,10,0,0q�5}     $
��Left=�5~     $	��Top=�5     �� 5�  
   -6   5        6 ��Name$6 
loginPanel��5�     $��83��5�     $��485,469,61,75q�5�     $y�Hidden=�5�     .+�5�     �� $	��110��6   $��20��5�     $��	129,1,0,0q�5�     $
��Left=�5�     $	��Top=�5�     
�������
���   �$��1q�5�     $
 5��5�     �� 5�     $��20��6   $��
129,5,38,0q�5�     $	��Top=�5�     
�������
���   �$��1q�5�     $
 5��5�     �� 5�     -7   5        $* login��5�     $	��154��5�     $��25��5�     $��64,8,0,0q�5�     $
��Left=�5�     $	��Top=�5�     $��Iniciar sesión��5�     ��5�     �� 5�     $	 0,0&�6*   $	 0,1&�6;   .��5�     � $ 0Q�6'   $ White��62   � 5�     $	 0.5Q�6'   $
 Gray��64   � 5�     $ 1Q�6'   $ White��62   5�     5�     �� 5�     $��82��6   $��20��5�     $��0q�5�     $
��Left=�5�     $	��Top=�5�     $��Center=�5�     $��Contraseña:��5�     $��5,0,5,5q�5�     �� 5�     -8   5        $
! user��5�     $	��110��5�     $��20��5�     $��129,-98,0,0q�5�     $
��Left=�5�     $	��Top=�5�     
������ ���	#000000007 �� SelectionBrush7 �	#003399FF5�     G� 5�     -9   5        8 G�Name$
8 pass��5�     $	��110��5�     $��20��5�     $��129,-73,0,0q�5�     $
��Left=�5�     $	��Top=�5�     
������ ���	#000000009 G� SelectionBrush9 �	#003399FF5�     �� 5�     $��25��6   $��29,-175,157,0q�5�     $��Center=�5�     $��Usuario:��5�     �� 5�     $��25��6   $��29,-125,157,0q�5�     $��Center=�5�     $��Contraseña:��5�     5�     �� 5�  
   -:   5        $6 loggedPanel��5�     $	��277��5�     $��85��5�     $��485,467,0,0q�5�     $
��Left=�5�     $	��Top=�5�     $y�Hidden=�5�     .+�5�     �� -;   5        $* logout��5�     $��77��5�     $��25��5�     $��	10,40,0,0q�5�     $��Center=�5�     $��Center=�5�     $��Salir��5�     ��5�     �� 5�     $	 0,0&�6*   $	 0,1&�6;   .��5�     � $ 0Q�6'   $ White��62   � 5�     $	 0.5Q�6'   $
 Gray��64   � 5�     $ 1Q�6'   $ White��62   5�     5�     �� 5�     -<   5        : ��Name$: welcomeText��5�     $	��277��5�     $��26��5�     $��	0,-80,0,0q�5�     $��Center=�5�     $��Center=�5�     $��Center=�5�     
������ 
�������; �� 
IsReadOnly; . 7 �	#003399FF��5�     �� 5�     $+ Absolute=�6*   $	 0,0&�6A   $
 0,20&�6R   .��5�     � $ 	#00ABADB3��6'   � 5�     $ 1Q�6'   $ 	#00E2E3EA��62   5�     . �5�     /� .T�5�     J� < �� TextAlignment$< Center=�6    .*�67   .*�5�     �� $�Bienvenido, ��6    �� 5�     $� ��6    �� 5�     -=   5        = ��Name$= username��5�     $
��Bold"�5�     $	�xxx��5�      �� 5�     $� ��6    �� 5�     $�!��6   5�     5�     5�     5�     5     5     