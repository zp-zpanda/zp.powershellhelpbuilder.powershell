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

            # Off-by-one. Last token is not in definition
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

            # Comment region must be within $Proximity lines
            $StartLine = $Tokens[$Index].Extent.StartLineNumber - $Proximity

            # Walk upwards and collect in reverse
            For ($I = ($Index - 1); $I -Ge 0; $I--)
            {
                $CurrentToken = $Tokens[$I]

                If ($CurrentToken.Extent.EndLineNumber -Lt $StartLine)
                {
                    # Walked past region already
                    Break
                }

                If ($CurrentToken.Kind -Eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)

                    # Extend comment region
                    $StartLine = $CurrentToken.Extent.StartLineNumber - 1
                }
                ElseIf ($CurrentToken.Kind -Ne [TokenKind]::NewLine)
                {
                    Break
                }
            }

            # Tokens were collected in reverse, so reverse again
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

            # We first assume comment region extends to end of file
            $EndLine = [int]::MaxValue

            # Walk downwards and collect
            For ($I = $StartIndex.Value; $I -Lt $Tokens.Length; $I++)
            {
                $CurrentToken = $Tokens[$I]

                If ($CurrentToken.Extent.StartLineNumber -Gt $EndLine)
                {
                    # Walked past region already
                    $StartIndex = $I
                    Break
                }

                If ($CurrentToken.Kind -Eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)

                    # Set comment region to end just after comment
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

        # No command, no tokens
        If ($Command.Length -Eq 0)
        {
            Return ,[Token[]]@()
        }

        $CommandInfo = Get-Command $Command

        # Resolve aliases
        If ($CommandInfo -Is [AliasInfo])
        {
            $CommandInfo = $CommandInfo.ResolvedCommand
        }

        # Tokens can only be in scripts and functions
        If (($CommandInfo -IsNot [FunctionInfo]) -And ($CommandInfo -IsNot [ExternalScriptInfo]))
        {
            Write-Error ("Command `'{0}`' is not a script or a function." -F $Command) -ErrorAction Stop
        }

        # Get ASTs
        $Ast = $CommandInfo.ScriptBlock.Ast
        $FunctionAst = $Ast -As [FunctionDefinitionAst]
        $RootAst = $Ast
        While ($null -Ne $RootAst.Parent)
        {
            $RootAst = $RootAst.Parent
        }

        # Get tokens
        $RootAstTokens = $RootAstErrors = $null
        [Parser]::ParseInput($RootAst.Extent.Text, [ref]$RootAstTokens, [ref]$RootAstErrors) | Out-Null
        
        $StartTokenIndex = $EndTokenIndex = $null

        If ($null -Ne $FunctionAst)
        {
            # Get tokens in front of function
            $FunctionDeclarationTokenIndex = Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex 0
            $PrecedingCommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $FunctionDeclarationTokenIndex
            
            If ($PrecedingCommentTokens.Count -Gt 0)
            {
                Return ,$PrecedingCommentTokens
            }

            # Tokens are in function
            $StartTokenIndex = (Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $FunctionAst.Body.Extent -StartIndex 0) + 1
            $EndTokenIndex = Get-LastTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex $StartTokenIndex
            
            # Sanity checks
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
            # Tokens must be in script
            $StartTokenIndex = 0
            $EndTokenIndex = $RootAstTokens.Length - 1
        }
        Else
        {
            <#
            Command is defined as scriptblock variable,
            which is rare in general but common when remoting.

            Example:

            $FooFunction = { ... }
            Set-Item Function:Foo $FooFunction
            Foo # Runs $FooFunction
            #>
            
            # Tokens must be in script block
            $StartTokenIndex = (Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex 0) + 1
            $EndTokenIndex = Get-LastTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex $StartTokenIndex
            
            # Sanity checks
            If ($RootAstTokens[$StartTokenIndex - 1].Kind -Ne [TokenKind]::LCurly)
            {
                Write-Error "Unexpected first token in script block." -ErrorAction Stop
            }

            If ($RootAstTokens[$EndTokenIndex].Kind -Ne [TokenKind]::RCurly)
            {
                Write-Error "Unexpected last token in script block." -ErrorAction Stop
            }
        }

        # Get tokens in definition
        While ($true)
        {
            # Get tokens at start
            $CommentTokens = Get-CommentTokens -Tokens $RootAstTokens -StartIndex ([ref]$StartTokenIndex)
            
            If ($CommentTokens.Count -Eq 0)
            {
                Break
            }

            If ($Ast -Eq $RootAst)
            {
                # Check if comments are close enough to be for first function instead of script
                $EndBlock = ([ScriptBlockAst]$Ast).EndBlock

                # Only unnamed end blocks can span the whole script
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

        # Get tokens at end
        $CommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $EndTokenIndex -Proximity $RootAstTokens[$EndTokenIndex].Extent.StartLineNumber
        If ($CommentTokens -Gt 0)
        {
            Return ,$CommentTokens
        }

        # No tokens found
        Return ,[Token[]]@()
    }
    End
    {}
}
Export-ModuleMember -Function @("Get-HelpCommentTokens")
