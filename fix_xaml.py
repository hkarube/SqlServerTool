import re

file = r'C:\dev\Project\SqlServerTool\Views\TableDetailWindow.xaml'
with open(file, encoding='utf-8') as f:
    text = f.read()

old = '''            <TextBlock Grid.Row="0" Text="■ 実行SQLログ（このセッション）"
                       FontSize="11" FontWeight="Bold" Foreground="DarkBlue"
                       Margin="6,3,0,1"/>
            <ListBox Grid.Row="1" x:Name="SqlLogList"
                     ItemsSource="{Binding SessionLog}"
                     FontFamily="Consolas" FontSize="11"
                     ScrollViewer.HorizontalScrollBarVisibility="Auto">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <TextBlock Text="{Binding Display}" TextWrapping="NoWrap"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>'''

new = '''            <Grid Grid.Row="0" Margin="0,2,0,1">
                <TextBlock Text="■ 実行SQLログ（このセッション）"
                           FontSize="11" FontWeight="Bold" Foreground="DarkBlue"
                           VerticalAlignment="Center" Margin="6,0,0,0"/>
                <Button Content="全履歴を開く" HorizontalAlignment="Right"
                        FontSize="10" Padding="6,1" Margin="0,0,4,0"
                        Click="OpenHistoryButton_Click"/>
            </Grid>
            <ListBox Grid.Row="1" x:Name="SqlLogList"
                     ItemsSource="{Binding SessionLog}"
                     FontFamily="Consolas" FontSize="11"
                     ScrollViewer.HorizontalScrollBarVisibility="Auto"
                     MouseDoubleClick="SqlLogList_MouseDoubleClick">
                <ListBox.ContextMenu>
                    <ContextMenu>
                        <MenuItem Header="SQLをコピー" Click="SqlLogCopy_Click"/>
                        <MenuItem Header="全文を表示" Click="SqlLogView_Click"/>
                    </ContextMenu>
                </ListBox.ContextMenu>
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <TextBlock Text="{Binding Display}" TextWrapping="NoWrap"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>'''

if old in text:
    text = text.replace(old, new, 1)
    with open(file, 'w', encoding='utf-8') as f:
        f.write(text)
    print('OK')
else:
    print('NOT FOUND')
