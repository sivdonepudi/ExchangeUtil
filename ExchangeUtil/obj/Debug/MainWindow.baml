   M S B A M L   `   `   `  ����G  CExchangeUtil, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:ExchangeUtilExchangeUtil  5   	       ExchangeUtil.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ,localclr-namespace:ExchangeUtil     5   	     � Title$  Exchange Util��5   	   $	��750��6   $
��1000��6,   
�������.��5	      � $��0,0,0,0q�6   .+�5
   
   �� $��Login��6   $
��Left=�6    $��	70,25,0,0q�6;   $	��Top=�6N   �� 5   
   -   5        $��TxtLogin��5      $
��Left=�6"   $��23��6=   $��
172,25,0,0q�6I    �� TextWrapping$
 Wrap=�6]   $�� ��6q   $	��Top=�6y   $	��350��6�   �� 5   
   $��Password��6   $
��Left=�6#   $��	70,56,0,0q�6>   $	��Top=�6Q   �� 5   
   -   5        $��TxtPassword��5      $
��Left=�6&   $��23��6A   $��
172,60,0,0q�6M   $
 Wrap=�6a   $�� ��6u   $	��Top=�6~   $	��350��6�   �� 5   
   $	��Url��6   $
��Left=�6   $��	71,97,0,0q�69   $	��Top=�6L    J� RenderTransformOrigin$ -1.331,0.105&�6d   �� 5   
   -   5        $��TxtUrl��5      $
��Left=�6    $��23��6;   $��172,101,0,0q�6G   $
 Wrap=�6\   $5��/https://outlook.office365.com/EWS/Exchange.asmx��6p   $	��Top=�6�   $	��350��6�   �� 5   
   $��BackDays��6   $
��Left=�6#   $��
72,136,0,0q�6>   $	��Top=�6R   $ 0.325,-1.381&�6j   �� 5   
   -   5        $��TxtBackDays��5      $
��Left=�6%   $��23��6@   $��171,139,0,0q�6L   $
 Wrap=�6a   $��0��6u   $	��Top=�6~   $	��120��6�   �� 5   
   $��Exchange Id��6   $
��Left=�6&   $��
72,170,0,0q�6A   $	��Top=�6U   $ 0.325,-1.381&�6m   �� 5   
   -   5         ��Name$ TxtUser��5      $
��Left=�6#   $��23��6>   $��171,174,0,0q�6J   $
 Wrap=�6_   $�� ��6s   $	��Top=�6{   $	��350��6�   �� 5   
   -   5        $��BtnGo��5      $��Go��6   $
��Left=�6+   $��553,173,0,0q�6F   $	��Top=�6[   $��75��6s   6~   �� 5   
   $��Stretch=�6   $��15��65   $��
10,216,0,0q�6A   $	��Top=�6U   �� 5   
   -   5        $��	BtnDelete��5   "   $��Delete��6   $��Right=�63   {�.  $��Center=�6a   $��
0,262,32,0q�6�   $	��Top=�6�   $��75��6�   6�   �� 5   
   -   5        $��LblInfo��5      $
��Left=�6   $��64,236,0,453q�6:   $	��Top=�6P   $
��Auto��6i   $��35��6v   6�   &   System.Windows.Controls.DataGrid  5   
   -	   5        $��DgData��5      $��Stretch=�6!      ColumnWidth5  /System.Windows.Controls.DataGridLengthConverter$ * 6?      CanUserAddRows .  $��
21,304,0,0q�6�   $��Stretch=�6�      AutoGenerateColumns .  ��6f   �� *	GridItems+   Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn   Header$ Subject��6%   1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 66   �� *	Subject+
   HeaderStyle
 5      �� 5       �� 
TargetType=  7System.Windows.Controls.Primitives.DataGridColumnHeader"  6    .�5      ��  �� Property � �� �� Value$ Center=�5   %   5      5!        5"      $ Start��6%   	 64   �� *	StartDate+
 5#      �� 5$      "  6    .�5%      ��  � ��$ Center=�5&   %   5'      5)        5*      $ 	Organizer��6%   	 68   �� *	Organizer+
 5+      �� 5,      "  6    .�5-      ��  � ��$ Center=�5.   %   5/      51        52      $ 
Categories��6%   	 69   �� *
Categories+
 53      �� 54      "  6    .�55      ��  � ��$ Center=�56   %   57      59        5:      $ UniqueId��6%      
Visibility$ Hidden=�6T   	 67   �� *
UniqueId+6T   4  .System.Windows.Controls.DataGridCheckBoxColumn  5;      	 6*   ��  �� Path$ 
IsSelected� 5<      �� 5=      -
   5        $��chkSelectAll��5=   4   $��Select��6#   6�   
 5?      �� 5@      "  6    .�5A      ��  � ��$ Center=�5B   %   5C      5E      5G      5J      5K      