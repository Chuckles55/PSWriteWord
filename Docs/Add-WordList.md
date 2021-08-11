---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordList

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordList [[-WordDocument] <Container>] [[-Paragraph] <InsertBeforeOrAfter>] [[-Type] <ListItemType>]
 [[-ListData] <Array>] [[-BehaviourOption] <Int32>] [[-ListLevels] <Array>] [[-Supress] <Boolean>]
 [<CommonParameters>]
```

## DESCRIPTION
Adds (sequentially) list items into an existing word-document-object from a supplied array. Optionally, for each item, a listlevel 
(i.e. hierarchical relationship relative to the first array member) can be specified. Lists can be either bullited or numbered.

## EXAMPLES

### Example 1
```powershell
PS C:\> Add-WordList -WordDocument "$env:TEMP\myword.docx" -type 'bulleted' -listdata @('a','b','c') -Supress:$true
```
Add a simple list, all at the same level, from a specified arra

### Example 2
```powershell
PS C:\> Add-WordList -WordDocument "$env:TEMP\myword.docx" -type 'bulleted' -listdata @('a','b','c') -listlevels @(1,2,3) -Supress:$true
```
Add a simple list, each member subordinate in level to its predecessor within the array

### Example 3
```powershell
PS C:\> Add-WordList -WordDocument "$env:TEMP\myword.docx" -type 'bulleted' -listdata @('a','b','c') -listlevels @(1,1,1) -Supress:$true
```
Add a simple list, all at the same level, but indented by one level (e.g. for readability)

## PARAMETERS

### -BehaviourOption
{{Fill BehaviourOption Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListData
{{Fill ListData Description}}

```yaml
Type: Array
Parameter Sets: (All)
Aliases: DataTable

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListLevels
A collection of values, mapped 1-to-1 to the specified listdata, determining the appropriae list indentation and symbols indicating such.

```yaml
Type: Array
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
{{Fill Paragraph Description}}

```yaml
Type: InsertBeforeOrAfter
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Supress
Switch. If TRUE, do NOT display resulting XML on screen.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
{{ Fill Type Description }}

```yaml
Type: ListItemType
Parameter Sets: (All)
Aliases: ListType
Accepted values: Bulleted, Numbered

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordDocument
.NET object (representing WORD document), typically created by 'new-worddocument', specifying the target of this command's actions

```yaml
Type: Container
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Container
InsertBeforeOrAfter

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
