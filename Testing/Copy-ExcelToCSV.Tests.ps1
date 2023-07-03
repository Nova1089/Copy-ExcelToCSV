<#
- Problem statement
- Use cases and features
- Minimum viable product
- Inputs
    - both input param and output param. Use both.
    - input param with no output param: Use input param with default output param.
    - output param with no input param: Prompt for input param and use provided output param.
    - neither an output param or input param: Prompt for input param and use default output param.
- Outputs
- Program flow
- Functions
- Classes and namespaces
- Input validation
    - param input path not an xlsx
    - param input path not a string
    - param input path empty
    - param input path missing extension
    - param input path not found
    - param output path not csv
    - param output path not a string
    - param output path is empty
    - param output path missing extension
    - param output path is in a folder that doesn't exist
    - all the previous stuff but for the prompted data
- Output validation
    - Has same number of records
    - Has all the same columns
- To Do
- Done but not tested
- Done and tested
    - Add message "Importing excel file..."
    - Add message "Exporting to CSV..."
#>

BeforeAll {
    # Optional
    # BeforeAll runs once at the beginning of the file.

    function Get-Functions($filePath)
    {
        $script = Get-Command $filePath
        return $script.ScriptBlock.AST.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
    }

    $scriptFolder = Split-Path -Path $PSScriptRoot -Parent
    $path = "$scriptFolder\Copy-ExcelToCsv.ps1"
    Get-Functions $path | Invoke-Expression

    $validExcelPath = "$PSScriptRoot\FruitCity.xlsx"
}

Describe "Function-Name" {
    BeforeEach {
        # Optional
        # Runs once before each test (It block) within the current Describe or Context block.
    }

    Context "When passing a something" {
        It "Should do/return something" {
            # Pipe values you want to test to Should
            # i.e: $result | Should -Contain $expected
            # More assertion examples: https://pester.dev/docs/assertions/
            # -Be, -Contain (value present in collection), -Not -Be, -BeExactly (tests for object equality), -BeGreaterThan, -BeGreaterOrEqual
            # -BeLessThan, -BeLessOrEqual, -BeIn (value is present in array/collection), -BeLike (wildcard pattern), -BeNullOrEmpty, -BeOfType
            # -HaveCount, -Match (regex comparison)
            # Mock behavior of existing function with an alternate implementation. Mock FunctionToMock { # alternate implementaton } 
            # Skipping: You can skip describe or context block with -skip operator. i.e. Describe "Some-Function" -Skip {
        }
    }

    AfterEach {
        # Optional
        # Runs once after each test (It block) within the current Describe or Context block.
    }
}

Describe "Validate-InputPath" {
    Context "When passed valid XLSX" {
        It "Should return true" {
            Validate-InputPath $validExcelPath | Should -Be $true
        }        
    }
    Context "When passed valid XLSX with double quotes" {
        It "Should return true" {
            Validate-InputPath "`"$validExcelPath`"" | Should -Be $true
        }
    }
    Context "When passed invalid path" {
        It "Should return false" {
            Validate-InputPath "InvalidPath" | Should -Be $false
        }
    }
    Context "When passed xlsx that doesn't exist" {
        It "Should return false" {
            Validate-InputPath "C:\DontExist.xlsx" | Should -Be $false
        }
    }
}

Describe "Validate-OutputPath" {
    Context "When passed valid path" {
        It "Should return true" {
            Validate-OutputPath "C:\ValidPath.csv" | Should -Be $true
        }        
    }
    Context "When passed invalid path" {
        It "Should return false" {
            Validate-OutputPath "Invalid Path" | Should -Be $false
        }
    }
}

AfterAll {
    # Optional
    # Runs once at the end of the file.
}