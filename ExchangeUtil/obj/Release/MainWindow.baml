   M S B A M L   `   `   `  ���� G  CExchangeUtil, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:ExchangeUtilExchangeUtil      ExchangeUtil.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ,localclr-namespace:ExchangeUtil       � Title$  Exchange Util��$	��750��$
��1000��
�������.��� $��0,0,0,0q�.+��� $��Login��$
��Left=�$��	70,25,0,0q�$	��Top=��� -   $��TxtLogin��$
��Left=�$��23��$��
172,25,0,0q� �� TextWrapping$
 Wrap=�$�� ��$	��Top=�$	��350���� $��Password��$
��Left=�$��	70,56,0,0q�$	��Top=��� -   $��TxtPassword��$
��Left=�$��23��$��
172,60,0,0q�$
 Wrap=�$�� ��$	��Top=�$	��350���� $	��Url��$
��Left=�$��	71,97,0,0q�$	��Top=� J� RenderTransformOrigin$ -1.331,0.105&��� -   $��TxtUrl��$
��Left=�$��23��$��172,101,0,0q�$
 Wrap=�$5��/https://outlook.office365.com/EWS/Exchange.asmx��$	��Top=�$	��350���� $��BackDays��$
��Left=�$��
72,136,0,0q�$	��Top=�$ 0.325,-1.381&��� -   $��TxtBackDays��$
��Left=�$��23��$��171,139,0,0q�$
 Wrap=�$��0��$	��Top=�$	��120���� $��Exchange Id��$
��Left=�$��
72,170,0,0q�$	��Top=�$ 0.325,-1.381&��� -    ��Name$ TxtUser��$
��Left=�$��23��$��171,174,0,0q�$
 Wrap=�$�� ��$	��Top=�$	��350���� -   $��BtnGo��$��Go��$
��Left=�$��553,173,0,0q�$	��Top=�$��75���� $��Stretch=�$��15��$��
10,216,0,0q�$	��Top=��� -   $��	BtnDelete��$��Delete��$��Right=�{�.  $��Center=�$��
0,262,32,0q�$	��Top=�$��75���� -   $��LblInfo��$
��Left=�$��64,236,0,453q�$	��Top=�$
��Auto��$��35��&   System.Windows.Controls.DataGrid  -	   $��DgData��$��Stretch=�   ColumnWidth5  /System.Windows.Controls.DataGridLengthConverter$ *    CanUserAddRows .  $��
21,304,0,0q�$��Stretch=�   AutoGenerateColumns .  ���� *	GridItems+   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Subject��1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 �� *	Subject+
   HeaderStyle
 ��  �� 
TargetType=  7System.Windows.Controls.Primitives.DataGridColumnHeader"  .���  �� Property � �� �� Value$ Center=�  $ Start��	 �� *	StartDate+
 �� "  .���  � ��$ Center=�  $ 	Organizer��	 �� *	Organizer+
 �� "  .���  � ��$ Center=�  $ 
Categories��	 �� *
Categories+
 �� "  .���  � ��$ Center=�  $ UniqueId��   
Visibility$ Hidden=�	 �� *
UniqueId+4  .System.Windows.Controls.DataGridCheckBoxColumn  	 ��  �� Path$ 
IsSelected� �� -
   $��chkSelectAll��$��Select��
 �� "  .���  � ��$ Center=�