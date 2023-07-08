using namespace System
using namespace System.Collections.Generic
using namespace System.Linq
using namespace System.Management.Automation
using namespace System.Management.Automation.Language

$script:CommentBlockProximity = 2

function Get-HelpCommentTokens
{
    [CmdletBinding()]
    [OutputType([Token[]])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]
        $Command
    )

    process
    {
        function Get-FirstTokenIndex
        {
            param
            (
                [Token[]]
                $Tokens,

                [IScriptExtent]
                $Extent,

                [int]
                $StartIndex
            )

            function IsBefore
            {
                param
                (
                    [IScriptExtent]
                    $First,

                    [IScriptExtent]
                    $Second
                )

                if ($First.EndLineNumber -lt $Second.StartLineNumber)
                {
                    return $true
                }

                if ($First.EndLineNumber -eq $Second.StartLineNumber)
                {
                    return $First.EndColumnNumber -le $Second.StartColumnNumber
                }

                return $false
            }

            for ($I = $StartIndex; $I -lt $Tokens.Length; $I++)
            {
                if (-not (IsBefore -First $Tokens[$I].Extent -Second $Extent))
                {
                    break
                }
            }

            return $I
        }

        function Get-LastTokenIndex
        {
            param
            (
                [Token[]]
                $Tokens,

                [IScriptExtent]
                $Extent,

                [int]
                $StartIndex
            )

            function IsAfter
            {
                param
                (
                    [IScriptExtent]
                    $First,

                    [IScriptExtent]
                    $Second
                )

                if ($First.StartLineNumber -gt $Second.EndLineNumber)
                {
                    return $true
                }

                if ($First.StartLineNumber -eq $Second.EndLineNumber)
                {
                    return $First.StartColumnNumber -ge $Second.EndColumnNumber
                }

                return $false
            }

            for ($I = $StartIndex; $I -lt $Tokens.Length; $I++)
            {
                if (IsAfter -First $Tokens[$I].Extent -Second $Extent)
                {
                    break
                }
            }

            # Off-by-one. Last token is not in definition
            return ($I - 1) 
        }

        function Get-PrecedingCommentTokens
        {
            param
            (
                [Token[]]
                $Tokens,

                [int]
                $Index,

                [int]
                $Proximity = $script:CommentBlockProximity
            )

            $CommentTokens = [List[Token]]::new()

            # Comment region must be within $Proximity lines
            $StartLine = $Tokens[$Index].Extent.StartLineNumber - $Proximity

            # Walk upwards and collect in reverse
            for ($I = ($Index - 1); $I -ge 0; $I--)
            {
                $CurrentToken = $Tokens[$I]

                if ($CurrentToken.Extent.EndLineNumber -lt $StartLine)
                {
                    # Walked past region already
                    break
                }

                if ($CurrentToken.Kind -eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)

                    # Extend comment region
                    $StartLine = $CurrentToken.Extent.StartLineNumber - 1
                }
                elseif ($CurrentToken.Kind -ne [TokenKind]::NewLine)
                {
                    break
                }
            }

            # Tokens were collected in reverse, so reverse again
            $CommentTokens.Reverse()

            return ,$CommentTokens.ToArray()
        }

        function Get-CommentTokens
        {
            param
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
            for ($I = $StartIndex.Value; $I -lt $Tokens.Length; $I++)
            {
                $CurrentToken = $Tokens[$I]

                if ($CurrentToken.Extent.StartLineNumber -gt $EndLine)
                {
                    # Walked past region already
                    $StartIndex.Value = $I
                    break
                }

                if ($CurrentToken.Kind -eq [TokenKind]::Comment)
                {
                    $CommentTokens.Add($CurrentToken)

                    # Set comment region to end just after comment
                    $EndLine = $CurrentToken.Extent.EndLineNumber + 1
                }
                elseif ($CurrentToken.Kind -ne [TokenKind]::NewLine)
                {
                    $StartIndex.Value = $I
                    break
                }
            }

            return ,$CommentTokens.ToArray()
        }

        # No command, no tokens
        if ($Command.Length -eq 0)
        {
            return ,[Token[]]@()
        }

        $CommandInfo = Get-Command $Command

        # Resolve aliases
        if ($CommandInfo -is [AliasInfo])
        {
            $CommandInfo = $CommandInfo.ResolvedCommand
        }

        # Tokens can only be in scripts and functions
        If (($CommandInfo -isnot [FunctionInfo]) -and ($CommandInfo -isnot [ExternalScriptInfo]))
        {
            Write-Error ("Command `'{0}`' is not a script or a function." -f $Command) -ErrorAction Stop
        }

        # Get ASTs
        $Ast = $CommandInfo.ScriptBlock.Ast
        $FunctionAst = $Ast -as [FunctionDefinitionAst]
        $RootAst = $Ast
        While ($null -ne $RootAst.Parent)
        {
            $RootAst = $RootAst.Parent
        }

        # Get tokens
        $RootAstTokens = $RootAstErrors = $null
        $null = [Parser]::ParseInput($RootAst.Extent.Text, [ref]$RootAstTokens, [ref]$RootAstErrors)
        
        $StartTokenIndex = $EndTokenIndex = $null

        If ($null -ne $FunctionAst)
        {
            # Get tokens in front of function
            $FunctionDeclarationTokenIndex = Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex 0
            $PrecedingCommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $FunctionDeclarationTokenIndex
            
            if ($PrecedingCommentTokens.Count -gt 0)
            {
                return ,$PrecedingCommentTokens
            }

            # Tokens are in function
            $StartTokenIndex = (Get-FirstTokenIndex -Tokens $RootAstTokens -Extent $FunctionAst.Body.Extent -StartIndex 0) + 1
            $EndTokenIndex = Get-LastTokenIndex -Tokens $RootAstTokens -Extent $Ast.Extent -StartIndex $StartTokenIndex
            
            # Sanity checks
            if ($RootAstTokens[$StartTokenIndex - 1].Kind -ne [TokenKind]::LCurly)
            {
                Write-Error "Unexpected first token in function." -ErrorAction Stop
            }

            if ($RootAstTokens[$EndTokenIndex].Kind -ne [TokenKind]::RCurly)
            {
                Write-Error "Unexpected last token in function." -ErrorAction Stop
            }
        }
        elseif ($Ast -eq $RootAst)
        {
            # Tokens must be in script
            $StartTokenIndex = 0
            $EndTokenIndex = $RootAstTokens.Length - 1
        }
        else
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
            if ($RootAstTokens[$StartTokenIndex - 1].Kind -ne [TokenKind]::LCurly)
            {
                Write-Error "Unexpected first token in script block." -ErrorAction Stop
            }

            if ($RootAstTokens[$EndTokenIndex].Kind -ne [TokenKind]::RCurly)
            {
                Write-Error "Unexpected last token in script block." -ErrorAction Stop
            }
        }

        # Get tokens in definition
        while ($true)
        {
            # Get tokens at start
            $CommentTokens = Get-CommentTokens -Tokens $RootAstTokens -StartIndex ([ref]$StartTokenIndex)
            
            if ($CommentTokens.Count -eq 0)
            {
                break
            }

            if ($Ast -eq $RootAst)
            {
                # Check if comments are close enough to be for first function instead of script

                # Only unnamed end blocks can span the whole script
                if (-not $Ast.EndBlock.Unnamed)
                {
                    return ,$CommentTokens
                }

                $FirstStatement = [Enumerable]::FirstOrDefault($Ast.EndBlock.Statements)

                if ($FirstStatement -is [FunctionDefinitionAst])
                {
                    $LinesBetween = $FirstStatement.Extent.StartLineNumber - ($CommentTokens[-1]).Extent.EndLineNumber
                    
                    if ($LinesBetween -gt $script:CommentBlockProximity)
                    {
                        return ,$CommentTokens
                    }

                    break
                }
            }

            return ,$CommentTokens
        }

        # Get tokens at end
        $CommentTokens = Get-PrecedingCommentTokens -Tokens $RootAstTokens -Index $EndTokenIndex -Proximity $RootAstTokens[$EndTokenIndex].Extent.StartLineNumber
        if ($CommentTokens.Count -gt 0)
        {
            return ,$CommentTokens
        }

        # No tokens found
        return ,[Token[]]@()
    }
}

function Get-TextFromCommentTokens
{
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Token[]]
        $Tokens
    )

    process
    {
        $CommentLines = [List[string]]::new()

        foreach ($Text in $Tokens.Text)
        {
            $I = 0
            if ($Text[0] -eq "<")
            {
                # Text is multi-line block comment, which starts with '<#' and '#>', so skip them
                $LineStart = 2
                for ($I = 2; $I -lt ($Text.Length - 2); $I++)
                {
                    # Check for newline
                    if ($Text[$I] -eq "`n")
                    {
                        # Linux newline, proceed
                        $CommentLines.Add($Text.Substring($LineStart, $I - $LineStart))
                        $LineStart = $I + 1
                    }
                    elseif ($Text[$I] -eq "`r")
                    {
                        # Start of Windows newline (CR), chop text, but check for LF for newline also
                        $CommentLines.Add($Text.Substring($LineStart, $I - $LineStart))

                        # No need to check length here as comment text has at least '#>' at end
                        if ($Text[$I + 1] -eq "`n")
                        {
                            $I++
                        }

                        $LineStart = $I + 1
                    }
                }
                $CommentLines.Add($Text.Substring($LineStart, $I - $LineStart))
            }
            else
            {
                # Skip all leading '#'s as it is common to use more than one '#'
                while ($I -lt $Text.Length -and $Text[$I] -eq "#")
                {
                    $I++
                }

                $CommentLines.Add($Text.Substring($I))
            }
        }
        return $CommentLines.ToArray() -join [Environment]::NewLine
    }
}

Export-ModuleMember -Function @("Get-HelpCommentTokens", "Get-TextFromCommentTokens")
