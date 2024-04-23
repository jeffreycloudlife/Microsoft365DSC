[CmdletBinding()]
param(
)
$M365DSCTestFolder = Join-Path -Path $PSScriptRoot `
                        -ChildPath '..\..\Unit' `
                        -Resolve
$CmdletModule = (Join-Path -Path $M365DSCTestFolder `
            -ChildPath '\Stubs\Microsoft365.psm1' `
            -Resolve)
$GenericStubPath = (Join-Path -Path $M365DSCTestFolder `
    -ChildPath '\Stubs\Generic.psm1' `
    -Resolve)
Import-Module -Name (Join-Path -Path $M365DSCTestFolder `
        -ChildPath '\UnitTestHelper.psm1' `
        -Resolve)

$Global:DscHelper = New-M365DscUnitTestHelper -StubModule $CmdletModule `
    -DscResource "AADAccessReview" -GenericStubModule $GenericStubPath
Describe -Name $Global:DscHelper.DescribeHeader -Fixture {
    InModuleScope -ModuleName $Global:DscHelper.ModuleName -ScriptBlock {
        Invoke-Command -ScriptBlock $Global:DscHelper.InitializeScript -NoNewScope
        BeforeAll {

            $secpasswd = ConvertTo-SecureString "f@kepassword1" -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential ('tenantadmin@mydomain.com', $secpasswd)

            Mock -CommandName Confirm-M365DSCDependencies -MockWith {
            }

            Mock -CommandName Get-PSSession -MockWith {
            }

            Mock -CommandName Remove-PSSession -MockWith {
            }

            Mock -CommandName Update-MgBetaAccessReview -MockWith {
            }

            Mock -CommandName New-MgBetaAccessReview -MockWith {
            }

            Mock -CommandName Remove-MgBetaAccessReview -MockWith {
            }

            Mock -CommandName New-M365DSCConnection -MockWith {
                return "Credentials"
            }

            # Mock Write-Host to hide output during the tests
            Mock -CommandName Write-Host -MockWith {
            }
            $Script:exportedInstances =$null
            $Script:ExportMode = $false
        }
        # Test contexts
        Context -Name "The AADAccessReview should exist but it DOES NOT" -Fixture {
            BeforeAll {
                $testParams = @{
                    BusinessFlowTemplateId = "FakeStringValue"
                    CreatedBy = (New-CimInstance -ClassName MSFT_MicrosoftGraphuserIdentity -Property @{
                        IpAddress = "FakeStringValue"
                        odataType = "#microsoft.graph.a"
                        UserPrincipalName = "FakeStringValue"
                        HomeTenantId = "FakeStringValue"
                        HomeTenantName = "FakeStringValue"
                    } -ClientOnly)
                    Description = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Id = "FakeStringValue"
                    ReviewedEntity = (New-CimInstance -ClassName MSFT_MicrosoftGraphidentity -Property @{
                    } -ClientOnly)
                    ReviewerType = "FakeStringValue"
                    Settings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewSettings -Property @{
                        DurationInDays = 25
                        odataType = "#microsoft.graph.b"
                        RemindersEnabled = $True
                        JustificationRequiredOnApproval = $True
                        AutoApplyReviewResultsEnabled = $True
                        RecurrenceSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewRecurrenceSettings -Property @{
                            RecurrenceCount = 25
                            DurationInDays = 25
                            RecurrenceType = "FakeStringValue"
                            RecurrenceEndType = "FakeStringValue"
                        } -ClientOnly)
                        AccessRecommendationsEnabled = $True
                        AutoReviewEnabled = $True
                        ActivityDurationInDays = 25
                        AutoReviewSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphautoReviewSettings -Property @{
                            NotReviewedResult = "FakeStringValue"
                        } -ClientOnly)
                        MailNotificationsEnabled = $True
                    } -ClientOnly)
                    StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Status = "FakeStringValue"
                    Ensure = "Present"
                    Credential = $Credential;
                }

                Mock -CommandName Get-MgBetaAccessReview -MockWith {
                    return $null
                }
            }
            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Absent'
            }
            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }
            It 'Should Create the group from the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName New-MgBetaAccessReview -Exactly 1
            }
        }

        Context -Name "The AADAccessReview exists but it SHOULD NOT" -Fixture {
            BeforeAll {
                $testParams = @{
                    BusinessFlowTemplateId = "FakeStringValue"
                    CreatedBy = (New-CimInstance -ClassName MSFT_MicrosoftGraphuserIdentity -Property @{
                        IpAddress = "FakeStringValue"
                        odataType = "#microsoft.graph.a"
                        UserPrincipalName = "FakeStringValue"
                        HomeTenantId = "FakeStringValue"
                        HomeTenantName = "FakeStringValue"
                    } -ClientOnly)
                    Description = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Id = "FakeStringValue"
                    ReviewedEntity = (New-CimInstance -ClassName MSFT_MicrosoftGraphidentity -Property @{
                    } -ClientOnly)
                    ReviewerType = "FakeStringValue"
                    Settings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewSettings -Property @{
                        DurationInDays = 25
                        odataType = "#microsoft.graph.b"
                        RemindersEnabled = $True
                        JustificationRequiredOnApproval = $True
                        AutoApplyReviewResultsEnabled = $True
                        RecurrenceSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewRecurrenceSettings -Property @{
                            RecurrenceCount = 25
                            DurationInDays = 25
                            RecurrenceType = "FakeStringValue"
                            RecurrenceEndType = "FakeStringValue"
                        } -ClientOnly)
                        AccessRecommendationsEnabled = $True
                        AutoReviewEnabled = $True
                        ActivityDurationInDays = 25
                        AutoReviewSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphautoReviewSettings -Property @{
                            NotReviewedResult = "FakeStringValue"
                        } -ClientOnly)
                        MailNotificationsEnabled = $True
                    } -ClientOnly)
                    StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Status = "FakeStringValue"
                    Ensure = 'Absent'
                    Credential = $Credential;
                }

                Mock -CommandName Get-MgBetaAccessReview -MockWith {
                    return @{
                        AdditionalProperties = @{
                            '@odata.type' = "#microsoft.graph.AccessReview"
                        }
                        BusinessFlowTemplateId = "FakeStringValue"
                        CreatedBy = @{
                            IpAddress = "FakeStringValue"
                            '@odata.type' = "#microsoft.graph.a"
                            UserPrincipalName = "FakeStringValue"
                            HomeTenantId = "FakeStringValue"
                            HomeTenantName = "FakeStringValue"
                        }
                        Description = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                        EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Id = "FakeStringValue"
                        ReviewedEntity = @{
                        }
                        ReviewerType = "FakeStringValue"
                        Settings = @{
                            DurationInDays = 25
                            RecurrenceSettings = @{
                                RecurrenceCount = 25
                                DurationInDays = 25
                                RecurrenceType = "FakeStringValue"
                                RecurrenceEndType = "FakeStringValue"
                            }
                            RemindersEnabled = $True
                            JustificationRequiredOnApproval = $True
                            AutoApplyReviewResultsEnabled = $True
                            '@odata.type' = "#microsoft.graph.b"
                            AccessRecommendationsEnabled = $True
                            AutoReviewEnabled = $True
                            ActivityDurationInDays = 25
                            AutoReviewSettings = @{
                                NotReviewedResult = "FakeStringValue"
                            }
                            MailNotificationsEnabled = $True
                        }
                        StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Status = "FakeStringValue"

                    }
                }
            }

            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Present'
            }

            It 'Should return true from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It 'Should Remove the group from the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName Remove-MgBetaAccessReview -Exactly 1
            }
        }
        Context -Name "The AADAccessReview Exists and Values are already in the desired state" -Fixture {
            BeforeAll {
                $testParams = @{
                    BusinessFlowTemplateId = "FakeStringValue"
                    CreatedBy = (New-CimInstance -ClassName MSFT_MicrosoftGraphuserIdentity -Property @{
                        IpAddress = "FakeStringValue"
                        odataType = "#microsoft.graph.a"
                        UserPrincipalName = "FakeStringValue"
                        HomeTenantId = "FakeStringValue"
                        HomeTenantName = "FakeStringValue"
                    } -ClientOnly)
                    Description = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Id = "FakeStringValue"
                    ReviewedEntity = (New-CimInstance -ClassName MSFT_MicrosoftGraphidentity -Property @{
                    } -ClientOnly)
                    ReviewerType = "FakeStringValue"
                    Settings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewSettings -Property @{
                        DurationInDays = 25
                        odataType = "#microsoft.graph.b"
                        RemindersEnabled = $True
                        JustificationRequiredOnApproval = $True
                        AutoApplyReviewResultsEnabled = $True
                        RecurrenceSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewRecurrenceSettings -Property @{
                            RecurrenceCount = 25
                            DurationInDays = 25
                            RecurrenceType = "FakeStringValue"
                            RecurrenceEndType = "FakeStringValue"
                        } -ClientOnly)
                        AccessRecommendationsEnabled = $True
                        AutoReviewEnabled = $True
                        ActivityDurationInDays = 25
                        AutoReviewSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphautoReviewSettings -Property @{
                            NotReviewedResult = "FakeStringValue"
                        } -ClientOnly)
                        MailNotificationsEnabled = $True
                    } -ClientOnly)
                    StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Status = "FakeStringValue"
                    Ensure = 'Present'
                    Credential = $Credential;
                }

                Mock -CommandName Get-MgBetaAccessReview -MockWith {
                    return @{
                        AdditionalProperties = @{
                            '@odata.type' = "#microsoft.graph.AccessReview"
                        }
                        BusinessFlowTemplateId = "FakeStringValue"
                        CreatedBy = @{
                            IpAddress = "FakeStringValue"
                            '@odata.type' = "#microsoft.graph.a"
                            UserPrincipalName = "FakeStringValue"
                            HomeTenantId = "FakeStringValue"
                            HomeTenantName = "FakeStringValue"
                        }
                        Description = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                        EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Id = "FakeStringValue"
                        ReviewedEntity = @{
                        }
                        ReviewerType = "FakeStringValue"
                        Settings = @{
                            DurationInDays = 25
                            RecurrenceSettings = @{
                                RecurrenceCount = 25
                                DurationInDays = 25
                                RecurrenceType = "FakeStringValue"
                                RecurrenceEndType = "FakeStringValue"
                            }
                            RemindersEnabled = $True
                            JustificationRequiredOnApproval = $True
                            AutoApplyReviewResultsEnabled = $True
                            '@odata.type' = "#microsoft.graph.b"
                            AccessRecommendationsEnabled = $True
                            AutoReviewEnabled = $True
                            ActivityDurationInDays = 25
                            AutoReviewSettings = @{
                                NotReviewedResult = "FakeStringValue"
                            }
                            MailNotificationsEnabled = $True
                        }
                        StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Status = "FakeStringValue"

                    }
                }
            }


            It 'Should return true from the Test method' {
                Test-TargetResource @testParams | Should -Be $true
            }
        }

        Context -Name "The AADAccessReview exists and values are NOT in the desired state" -Fixture {
            BeforeAll {
                $testParams = @{
                    BusinessFlowTemplateId = "FakeStringValue"
                    CreatedBy = (New-CimInstance -ClassName MSFT_MicrosoftGraphuserIdentity -Property @{
                        IpAddress = "FakeStringValue"
                        odataType = "#microsoft.graph.a"
                        UserPrincipalName = "FakeStringValue"
                        HomeTenantId = "FakeStringValue"
                        HomeTenantName = "FakeStringValue"
                    } -ClientOnly)
                    Description = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Id = "FakeStringValue"
                    ReviewedEntity = (New-CimInstance -ClassName MSFT_MicrosoftGraphidentity -Property @{
                    } -ClientOnly)
                    ReviewerType = "FakeStringValue"
                    Settings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewSettings -Property @{
                        DurationInDays = 25
                        odataType = "#microsoft.graph.b"
                        RemindersEnabled = $True
                        JustificationRequiredOnApproval = $True
                        AutoApplyReviewResultsEnabled = $True
                        RecurrenceSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphaccessReviewRecurrenceSettings -Property @{
                            RecurrenceCount = 25
                            DurationInDays = 25
                            RecurrenceType = "FakeStringValue"
                            RecurrenceEndType = "FakeStringValue"
                        } -ClientOnly)
                        AccessRecommendationsEnabled = $True
                        AutoReviewEnabled = $True
                        ActivityDurationInDays = 25
                        AutoReviewSettings = (New-CimInstance -ClassName MSFT_MicrosoftGraphautoReviewSettings -Property @{
                            NotReviewedResult = "FakeStringValue"
                        } -ClientOnly)
                        MailNotificationsEnabled = $True
                    } -ClientOnly)
                    StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                    Status = "FakeStringValue"
                    Ensure = 'Present'
                    Credential = $Credential;
                }

                Mock -CommandName Get-MgBetaAccessReview -MockWith {
                    return @{
                        BusinessFlowTemplateId = "FakeStringValue"
                        CreatedBy = @{
                            IpAddress = "FakeStringValue"
                            '@odata.type' = "#microsoft.graph.a"
                            UserPrincipalName = "FakeStringValue"
                            HomeTenantId = "FakeStringValue"
                            HomeTenantName = "FakeStringValue"
                        }
                        Description = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                        EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Id = "FakeStringValue"
                        ReviewedEntity = @{
                        }
                        ReviewerType = "FakeStringValue"
                        Settings = @{
                            AutoReviewSettings = @{
                                NotReviewedResult = "FakeStringValue"
                            }
                            '@odata.type' = "#microsoft.graph.b"
                            ActivityDurationInDays = 7
                            RecurrenceSettings = @{
                                RecurrenceCount = 7
                                DurationInDays = 7
                                RecurrenceType = "FakeStringValue"
                                RecurrenceEndType = "FakeStringValue"
                            }
                            DurationInDays = 7
                        }
                        StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Status = "FakeStringValue"
                    }
                }
            }

            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Present'
            }

            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It 'Should call the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName Update-MgBetaAccessReview -Exactly 1
            }
        }

        Context -Name 'ReverseDSC Tests' -Fixture {
            BeforeAll {
                $Global:CurrentModeIsExport = $true
                $Global:PartialExportFileName = "$(New-Guid).partial.ps1"
                $testParams = @{
                    Credential = $Credential
                }

                Mock -CommandName Get-MgBetaAccessReview -MockWith {
                    return @{
                        AdditionalProperties = @{
                            '@odata.type' = "#microsoft.graph.AccessReview"
                        }
                        BusinessFlowTemplateId = "FakeStringValue"
                        CreatedBy = @{
                            IpAddress = "FakeStringValue"
                            '@odata.type' = "#microsoft.graph.a"
                            UserPrincipalName = "FakeStringValue"
                            HomeTenantId = "FakeStringValue"
                            HomeTenantName = "FakeStringValue"
                        }
                        Description = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                        EndDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Id = "FakeStringValue"
                        ReviewedEntity = @{
                        }
                        ReviewerType = "FakeStringValue"
                        Settings = @{
                            DurationInDays = 25
                            RecurrenceSettings = @{
                                RecurrenceCount = 25
                                DurationInDays = 25
                                RecurrenceType = "FakeStringValue"
                                RecurrenceEndType = "FakeStringValue"
                            }
                            RemindersEnabled = $True
                            JustificationRequiredOnApproval = $True
                            AutoApplyReviewResultsEnabled = $True
                            '@odata.type' = "#microsoft.graph.b"
                            AccessRecommendationsEnabled = $True
                            AutoReviewEnabled = $True
                            ActivityDurationInDays = 25
                            AutoReviewSettings = @{
                                NotReviewedResult = "FakeStringValue"
                            }
                            MailNotificationsEnabled = $True
                        }
                        StartDateTime = "2023-01-01T00:00:00.0000000+01:00"
                        Status = "FakeStringValue"

                    }
                }
            }
            It 'Should Reverse Engineer resource from the Export method' {
                $result = Export-TargetResource @testParams
                $result | Should -Not -BeNullOrEmpty
            }
        }
    }
}

Invoke-Command -ScriptBlock $Global:DscHelper.CleanupScript -NoNewScope
