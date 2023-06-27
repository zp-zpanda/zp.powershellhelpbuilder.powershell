Using Namespace System.Management.Automation
Using Namespace System.Management.Automation.Language
$Script:CommentBlockProximity = 2
Function Get-HelpCommentTokens
{
    [CmdletBinding()]
    [OutputType([Token[]])]
    Param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]
        $Command
    )
    Begin
    {}
    Process
    {
        Function Get-FirstTokenIndex
        {
            Param
            (
                [Token[]]
                $Tokens,
                [IScriptExtent]
                $Extent,
                [int]
                $StartIndex
            )
            Function IsBefore
            {
                Param
                (
                    [IScriptExtent]
                    $First,
                    [IScriptExtent]
                    $Second
                )
                If ($First.EndLineNumber -Lt $Second.StartLineNumber)
                {
                    Return $true
                }
                If ($First.EndLineNumber -Eq $Second.StartLineNumber)
                {
                    Return $First.EndColumnNumber -Le $Second.StartColumnNumber
                }
                Return $false
            }
            For ($I = $StartIndex; $I -Lt $Tokens.Length; ++$I)
            {
                If (-Not (IsBefore -First $Tokens[$I].Extent -Second $Extent))
                {
                    Break
                }
            }
            Return $I
        }
        Function Get-LastTokenIndex
        {
            Param
            (
                [Token[]]
                $Tokens,
                [IScriptExtent]
                $Extent,
                [int]
                $StartIndex
            )
            Function IsAfter
            {
                Param
                (
                    [IScriptExtent]
                    $First,
                    [IScriptExtent]
                    $Second
                )
                If ($First.StartLineNumber -Gt $Second.EndLineNumber)
                {
                    Return $true
                }
                If ($First.StartLineNumber -Eq $Second.EndLineNumber)
                {
                    Return $First.StartColumnNumber -Ge $Second.EndColumnNumber
                }
                Return $false
            }
            For ($I = $StartIndex; $I -Lt $Tokens.Length; ++$I)
            {
                If (IsAfter -First $Tokens[$I].Extent -Second $Extent)
                {
                    Break
                }
            }
            Return ($I - 1)
        }
        Function Get-PrecedingCommentTokens
        {
            Param
            (
                [Token[]]
                $Tokens,
                [int]
                $Index,
                [int]
                $Proximity = $Script:CommentBlockProximity
            )
            $CommentTokens = [List[Token]]::new()
            $StartLine = $Tokens[$Index].Extent.StartLineNumber - $Proximity
            For ($I = ($Index - 1); $I -Ge 0; $I--)
            {
                $CurrentToken = $Tokens[$I]
                If ($CurrentToken.Extent.EndLineNumber -Lt $StartLine)
                {
                    Break
                }
                If ($CurrentToken.Kind -Eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)
                    $StartLine = $CurrentToken.Extent.StartLineNumber - 1
                }
                ElseIf ($CurrentToken.Kind -Ne [TokenKind]::NewLine)
                {
                    Break
                }
            }
            $CommentTokens.Reverse()
            Return ,$CommentTokens.ToArray()
        }
        Function Get-CommentTokens
        {
            Param
            (
                [Token[]]
                $Tokens,
                [ref]
                $StartIndex
            )
            $CommentTokens = [List[Token]]::new()
            $EndLine = [int]::MaxValue
            For ($I = $StartIndex.Value; $I -Lt $Tokens.Length; $I++)
            {
                $CurrentToken = $Tokens[$I]
                If ($CurrentToken.Extent.StartLineNumber -Gt $EndLine)
                {
                    $StartIndex = $I
                    Break
                }
                If ($CurrentToken.Kind -Eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)
                    $EndLine = $CurrentToken.Extent.EndLineNumber + 1
                }
                ElseIf ($CurrentToken.Kind -Ne [TokenKind]::NewLine)
                {
                    $StartIndex = $I
                    Break
                }
            }
            Return ,$CommentTokens.ToArray()
        }
        If ($Command.Length -Eq 0)
        {
            Return ,[Token[]]@()
        }
        $CommandInfo = Get-Command $Command
        If ($CommandInfo -Is [AliasInfo])
        {
            $CommandInfo = $CommandInfo.ResolvedCommand
        }
        If (($CommandInfo -IsNot [FunctionInfo]) -And ($CommandInfo -IsNot [ExternalScriptInfo]))
        {
            Write-Error "Command is not a script or a function." -ErrorAction Stop
        }
        $Ast = $CommandInfo.ScriptBlock.Ast
        $FunctionAst = $Ast -As [FunctionDefinitionAst]
        $RootAst = $Ast
        While ($null -Ne $RootAst.Parent)
        {
            $RootAst = $RootAst.Parent
        }
        $RootAstTokens = $RootAstErrors = $null
        [Parser]::ParseInput($RootAst.Extent.Text, [ref]$RootAstTokens, [ref]$RootAstErrors) | Out-Null
        $StartTokenIndex = $EndTokenIndex = $null
        If ($null -Ne $FunctionAst)
        {
            $FunctionDeclarationTokenIndex = Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex 0
            $PrecedingCommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $FunctionDeclarationTokenIndex
            If ($PrecedingCommentTokens.Count -Gt 0)
            {
                Return ,$PrecedingCommentTokens
            }
            $StartTokenIndex = (Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $FunctionAst.Body.Extent -StartIndex 0) + 1
            $EndTokenIndex = Get-LastTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex $StartTokenIndex
            If ($RootAstTokens[$StartTokenIndex - 1].Kind -Ne [TokenKind]::LCurly)
            {
                Write-Error "Unexpected first token in function." -ErrorAction Stop
            }
            If ($RootAstTokens[$EndTokenIndex].Kind -Ne [TokenKind]::RCurly)
            {
                Write-Error "Unexpected last token in function." -ErrorAction Stop
            }
        }
        ElseIf ($Ast -Eq $RootAst)
        {
            $StartTokenIndex = 0
            $EndTokenIndex = $RootAstTokens.Length - 1
        }
        Else
        {
            $StartTokenIndex = (Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex 0) + 1
            $EndTokenIndex = Get-LastTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex $StartTokenIndex
            If ($RootAstTokens[$StartTokenIndex - 1].Kind -Ne [TokenKind]::LCurly)
            {
                Write-Error "Unexpected first token in script block." -ErrorAction Stop
            }
            If ($RootAstTokens[$EndTokenIndex].Kind -Ne [TokenKind]::RCurly)
            {
                Write-Error "Unexpected last token in script block." -ErrorAction Stop
            }
        }
        While ($true)
        {
            $CommentTokens = Get-CommentTokens -Tokens $RootAstTokens -StartIndex ([ref]$StartTokenIndex)
            If ($CommentTokens.Count -Eq 0)
            {
                Break
            }
            If ($Ast -Eq $RootAst)
            {
                $EndBlock = ([ScriptBlockAst]$Ast).EndBlock
                If ($null -Eq $EndBlock -Or -Not $EndBlock.Unnamed)
                {
                    Return ,$CommentTokens
                }
                $FirstStatement = [Enumerable]::FirstOrDefault($EndBlock.Statements)
                If ($FirstStatement -Is [FunctionDefinitionAst])
                {
                    $LinesBetween = $FirstStatement.Extent.StartLineNumber - ($CommentTokens[-1]).Extent.EndLineNumber
                    If ($LinesBetween -Gt $Script:CommentBlockProximity)
                    {
                        Return ,$CommentTokens
                    }
                    Break
                }
            }
            Return ,$CommentTokens
        }
        $CommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $EndTokenIndex -Proximity $RootAstTokens[$EndTokenIndex].Extent.StartLineNumber
        If ($CommentTokens -Gt 0)
        {
            Return ,$CommentTokens
        }
        Return ,[Token[]]@()
    }
    End
    {}
}
Export-ModuleMember -Function @("Get-HelpCommentTokens")
