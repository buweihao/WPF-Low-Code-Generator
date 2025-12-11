import { SheetData } from '../../types';
import { isArrayType } from '../utils';

export const generatePLCPointXaml = (sheets: SheetData[], maxModules: number): string => {
  const sb: string[] = [];
  sb.push(`<UserControl x:Class="Core.PLCPoint"`);
  sb.push(`             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"`);
  sb.push(`             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"`);
  sb.push(`             xmlns:local="clr-namespace:Core">`);
  sb.push(`    <UserControl.Resources>`);
  sb.push(`        <local:ArrayToStringConverter x:Key="ArrayConv"/>`);
  sb.push(`    </UserControl.Resources>`);
  sb.push(`    <UserControl.DataContext>`);
  sb.push(`        <x:Static Member="local:PLCPointProperty.Instance"/>`);
  sb.push(`    </UserControl.DataContext>`);
  sb.push(`    <Grid>`);
  sb.push(`        <Grid.RowDefinitions>`);
  sb.push(`            <RowDefinition Height="Auto"/>`);
  sb.push(`            <RowDefinition Height="*"/>`);
  sb.push(`        </Grid.RowDefinitions>`);
  
  // Header: Frequency + Module Selector + Recorder
  sb.push(`        <StackPanel Orientation="Horizontal" Margin="5">`);
  sb.push(`            <TextBlock Text="Frequency (ms): " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding Frequency}" Width="80" Margin="0,0,20,0">`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">0</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">500</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">1000</sys:Int32>`);
  sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">3000</sys:Int32>`);
  sb.push(`            </ComboBox>`);
  
  sb.push(`            <TextBlock Text="Current Module: " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding CurrentModuleIndex}" Width="80" Margin="0,0,20,0">`);
  for(let i=1; i<=maxModules; i++) {
     sb.push(`                <sys:Int32 xmlns:sys="clr-namespace:System;assembly=mscorlib">${i}</sys:Int32>`);
  }
  sb.push(`            </ComboBox>`);

  sb.push(`            <TextBlock Text="Record: " VerticalAlignment="Center" Margin="0,0,5,0"/>`);
  sb.push(`            <ComboBox SelectedItem="{Binding RecordState}" Width="100">`);
  sb.push(`                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">Stopped</sys:String>`);
  sb.push(`                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">Recording</sys:String>`);
  sb.push(`            </ComboBox>`);

  sb.push(`        </StackPanel>`);

  sb.push(`        <ScrollViewer Grid.Row="1">`);
  sb.push(`            <StackPanel>`);
  
  // Single View: Displays CurrentModule Properties
  sb.push(`                <GroupBox Header="Current Module View" Margin="5">`);
  sb.push(`                    <StackPanel>`);
  
  sheets.forEach(sheet => {
    sb.push(`                        <Expander Header="${sheet.name}" Margin="2" IsExpanded="True">`);
    sb.push(`                            <StackPanel>`);
    
    sheet.points.forEach(p => {
      sb.push(`                                <Grid Margin="2">`);
      sb.push(`                                    <Grid.ColumnDefinitions>`);
      sb.push(`                                        <ColumnDefinition Width="200"/>`);
      sb.push(`                                        <ColumnDefinition Width="*"/>`);
      sb.push(`                                    </Grid.ColumnDefinitions>`);
      sb.push(`                                    <TextBlock Text="${p.KeyName} (${p.PropertyName}):" VerticalAlignment="Center"/>`);
      // Bind to CurrentX
      if (isArrayType(p.Type)) {
          sb.push(`                                    <TextBox Grid.Column="1" Text="{Binding Current${p.PropertyName}, Converter={StaticResource ArrayConv}}" IsReadOnly="True"/>`);
      } else {
          sb.push(`                                    <TextBox Grid.Column="1" Text="{Binding Current${p.PropertyName}}" IsReadOnly="True"/>`);
      }
      sb.push(`                                </Grid>`);
    });

    sb.push(`                            </StackPanel>`);
    sb.push(`                        </Expander>`);
  });

  sb.push(`                    </StackPanel>`);
  sb.push(`                </GroupBox>`);

  sb.push(`            </StackPanel>`);
  sb.push(`        </ScrollViewer>`);
  sb.push(`    </Grid>`);
  sb.push(`</UserControl>`);
  return sb.join('\n');
};
