   M S B A M L   `   `   `  ����R  NFilesAndFolderRetriever, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:FilesAndFolderRetrieverFilesAndFolderRetriever  5   	   (    "FilesAndFolderRetriever.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   7local%clr-namespace:FilesAndFolderRetriever     5   	     � Title$  Files and folders indexer��5   	   $	��219��6+   $	��654��68    � WindowStartupLocation$ CenterScreen=�6D    � 
ResizeMode$ NoResize=�6i   .��5	      � .+�5
   
   �� -   5         ��Name$ btnFolderBrowse��5
      $��Browse��6*   $
��Left=�6;   $��30��6V   $��
526,22,0,0q�6b   $	��Top=�6v   $	��102��6�   6�   �� 5   
   -   5         ��Name$ lblFolderPath��5      $��Folder path��6'   $
��Left=�6=   $��25��6X   $��	18,24,0,0q�6d   $	��Top=�6w   $��79��6�   $
��Left=�6�   $
��Bold"�6�   �� 5   
   -   5         ��Name$ txtPath��5      $
��Left=�6#   $��25��6>   $��
120,24,0,0q�6J    �� TextWrapping$
 Wrap=�6^   $	��Top=�6r   $	��376��6�   �� 5   
   -   5        $ 	btnSaveAs��5      $��Save As��6$   $
��Left=�66   $��30��6Q   $��
526,79,0,0q�6]   $	��Top=�6q   $	��102��6�   6�   �� 5   
   -   5        $ 	lblSaveTo��5      $��Save to��6#   $
��Left=�65   $��25��6P   $��	18,81,0,0q�6\   $	��Top=�6o   $��79��6�   $
��Left=�6�   $
��Bold"�6�   �� 5   
   -   5        $ txtSaveToPath��5      $
��Left=�6)   $��25��6D   $��
120,81,0,0q�6P   $
 Wrap=�6d   $	��Top=�6x   $	��376��6�   �� 5   
   -   5        $ btnStart��5      $��Start��6#   $
��Left=�63   $��30��6N   $��526,139,0,0q�6Z   $	��Top=�6o   $	��102��6�   $
��Bold"�6�   � 5   
   -   5         �Name$ prgBar��5      $
��Left=�6&   $��21��6A   $	��Top=�6M   $	��376��6e   $��120,139,0,0q�6q    � SmallChange$ 1Q�6�   5      5      