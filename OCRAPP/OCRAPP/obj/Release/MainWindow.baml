   M S B A M L   `   `   `  ���� A  =OCRAPP, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:OCRAPPOCRAPP      OCRAPP.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     &localclr-namespace:OCRAPP     -     �Name$  ocrmain�� � Title$# Optical Character Recognition��$	��914��$
��1619�� � 
ResizeMode$ NoResize=����	#FFFDFDFD.��� $��0,0,-714,-1563q� � ColumnDefinitions �� $
��640* ��� $	��77* ��� $	��56* ��� $	��38* ��� $	��15* ��� $
��346* ��� $��1155* �.+��� $#��Optical Character Recognition��$
��Left=�$��43��$	��Top=�$	��315��$��Tahoma)�$��20(�$
��Bold"�$��42,0,0,0q����	#FF8F8A8A$��5��$��1���� $
��Left=�$	��779��$��	10,74,0,0q�
���   �$	��Top=�$	��383���� $
��Left=�$	��779��$��
398,74,0,0q�
���   �$	��Top=�$	��534��$��6���� $
��Left=�$	��779��$��
113,74,0,0q�
���   �$	��Top=�$	��238��$��5��$��2���� $��
File Path:��$
��Left=�$��28��$��
20,114,0,0q�$	��Top=�$��60���� -    ��Name$ filepathtxt��$
��Left=�$��20��$��
85,114,0,0q� �� TextWrapping$
 Wrap=�$��No file chosen��$	��Top=�$	��201���� -    ��Name$ choosefilebtn��$��Choose File��$
��Left=�$��20��$��291,114,0,0q�$	��Top=�$��92���� -    ��Name$ 
inputimage��$
��Left=�$	��349��$��
20,179,0,0q�$	��Top=�$	��363��$��back.jpg��$
��Fill=��� -   $ 	submitbtn��$��Upload��$
��Left=�$��22��$��
20,535,0,0q�$	��Top=�$	��115���� -   $ clearbtn��$��Clear��$
��Left=�$��21��$��140,536,0,0q�$	��Top=�$	��113���� -   $ exitbtn��$
��Exit��$
��Left=�$��20��$��146,820,0,0q�$	��Top=�$	��172��$��5���� $��Editable Text��$
��Left=�$��28��$��
410,82,0,0q�$	��Top=�$	��159��$��14(�$
��Bold"��� $��Total Number of Words:��$
��Left=�$��28��$��
20,619,0,0q�$	��Top=�$	��137���� $��Most Occuring Word:��$
��Left=�$��28��$��
20,652,0,0q�$	��Top=�$	��137���� -   $ noofwordstxt��$
��Left=�$��23��$��162,624,0,0q�$
 Wrap=�$	��Top=�$	��221�����	#FFE6E5E5�� -	    ��Name$ resulttxtblk��$
��Left=�$	��678��$��410,115,0,0q�$
��Wrap=�$	��Top=�$	��514�����	#FFEEEEEE$��6���� -
   $ 	statustxt��$
��Left=�$��25��$��
20,147,0,0q�$
 Wrap=�$��Select Assignment Image��$	��Top=�$	��363�����	#FFF30909��Z� �� -   $ meanconfidencetxt��$
��Left=�$��23��$��162,596,0,0q�$
 Wrap=�$	��Top=�$	��221�����	#FFE6E5E5�� $��Mean Confidence:��$
��Left=�$��28��$��
20,591,0,0q�$	��Top=�$	��109���� $��Statistical Results��$
��Left=�$��28��$��
20,557,0,0q�$	��Top=�$	��363��$��14(�$
��Bold"��� -   	 ��Name$	 wordresulttxt��$
��Left=�$	��160��$��
20,682,0,0q�$	��Top=�$	��363��
 �� View
 �� .P��� �� -   $ saveastxtbtn��$��Save as txt��$
��Left=�$��23��$��410,798,0,0q�$	��Top=�$	��255��$��2��&   System.Windows.Controls.Calendar  -     Name$ remindercalender��$��5��$
��Left=�$��127,570,0,0q�$	��Top=�$	��166��$	��210���� -   $ 
duedatetxt��$
��Left=�$��23��$��134,198,0,0q�$
 Wrap=�$	��Top=�$	��200�����	#FFE6E5E5$��5���� $��	Due Date:��$��5��$
��Left=�$��131,173,0,0q�$	��Top=�$��26��$��62���� $��Upload Date:��$��5��$
��Left=�$��124,122,0,0q�$	��Top=�$��26��$��79���� -   $ uploaddatetxt��$
��Left=�$��23��$��129,147,0,0q�$
 Wrap=�$	��Top=�$	��205�����	#FFE6E5E5$��5���� $��Set Reminder��$
��Left=�$��28��$��
124,83,0,0q�$	��Top=�$	��159��$��14(�$
��Bold"� J� RenderTransformOrigin$ 1.975,1.071&�$��5���� $��5��$
��Left=�$	��176��$��146,596,0,0q�$	��Top=�$	��167���� -   $ setreminderbtn��$��Set Reminder��$��5��$
��Left=�$��65��$��131,232,0,0q�$	��Top=�$	��203���� -   $ dataset��$��5��$
��Left=�$	��219��$��134,312,0,0q�$
 Wrap=�$	��Top=�$	��200���� $
��Left=�$	��779��$��9,74,0,0q�
���   �$	��Top=�$	��423��$��6���� -   $ reminder0ffbtn��$��Reminder Off��$��5��$
��Left=�$��59��$��146,749,0,0q�$	��Top=�$	��172���� $��List of Assignments��$
��Left=�$��28��$��	24,82,0,0q�$	��Top=�$	��159��$��14(�$
��Bold"�$ 1.975,1.071&�$��6���� $��Upload Assignment��$
��Left=�$��28��$��	20,81,0,0q�$	��Top=�$	��159��$��14(�$
��Bold"��� -   $ datastatustxt��$
��Left=�$��25��$��129,541,0,0q�$
 Wrap=�$	��Top=�$	��205�����	#FFF30909$��5����Z� �� -   $ liststatustxt��$
��Left=�$��25��$��
24,110,0,0q�$
 Wrap=�$	��Top=�$	��398�����	#FFF30909$��6����Z� �� -   $	 listtxtlist��$��6��$
��Left=�$	��698��$��
24,142,0,0q�$	��Top=�$	��398��
 �� .P��� �� -   $ savestatustxt��$
��Left=�$��25��$��410,821,0,0q�$
 Wrap=�$	��Top=�$	��511�����	#FFF30909$��6����Z� �� -   $ saveasdocxbtn��$��5��$��Save as Docx��$��1��$
��Left=�$��
30,798,0,0q�$	��Top=�$	��251��$��23��