   M S B A M L   `   `   `  ����H  DDataValidator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:DataValidatorDataValidator  5   	       DataValidator.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:DataValidator     5   	     � Title$  
MainWindow��5   	   $	��350��6   $	��525��6)    � DataContext 5	      &    DataValidator.ValidatorViewModel  5
   
   .��5      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   .+�5   
   �� $��0��6    �� Orientation$ 
Horizontal=�6"   
�������.+�5      �� $��10q�6   $��80��6!   ��6,   ��  �� Path$ ImportEpmReportCommand�.��6\   
Import EPM6h   �� 5      $��10q�6   $��80��6!   ��6,   �� $ ImportSCCMReportCommand�.��6]   Import SCCM6j   �� 5      $��10q�6   $��80��6!   ��6,   �� $ ExportReportCommand�.��6Y   Export6a   5      �� 5   
   -   5         ��Name$
 atms��5      $��1��6!    �� SelectionMode$ Single=�6.   ��. $
��Auto=�6�   ��6E   �� *DifferentAtms+ �� Mode$ OneWay=� �� UpdateSourceTrigger$ PropertyChanged=�	 �� View	 5      �� 5      .P�5      �� $��Customer��6%   
 �� Width$	
 100��67    �� DisplayMemberBinding 6C   �� *
Customer+�� 5      $
��Name��6%   $
 50��63    6>   �� *Name+�� 5      $��SWDMUP��6%   $	
 100��65    6A   �� *SWDMUP+�� 5      $	��MUP��6%   $	
 100��62    6>   �� *MUP+5      5!      5"      5#      