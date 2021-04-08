
# =================================================================================
# Information Leadership Office365 Retention label and policy checking script
# =================================================================================
# NOTES/PREREQUISITES:
# Please ensure that when running this script you have the Exchange administrator, 
# SharePoint Administrator and Compliance Administrator role assigned to you 
# within your tenancy. 
#
# If you are required to activate the role via PIM (Privileged Identity Management), 
# make sure that this is done prior to running the script.
#
# This script also relies on the ImportExcel Module, please use the command:
# Install-Module ImportExcel in an elevated PowerShell session to install
# ---------------------------------------------------------------------------------

Import-Module ExchangeOnlineManagement
Import-Module ImportExcel
# =================================================================================
# Input 365 Admin account credentials
# =================================================================================

$AdminUsername = Read-Host -Prompt "Azure/Office 365 Admin User Account"

$AdminPassword = Read-Host -Prompt "Password" -AsSecureString
$adminCredentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminUsername, $AdminPassword

# =================================================================================
# Connect to SharePoint Online using MFA if required
# =================================================================================
connect-sposervice -Credential $adminCredentials

# =================================================================================
# Connect to Remote IPPS Session
# Note: This will require WinRM Basic to be allowed on your machine, if this is 
# blocked by GPO please contact your administrator or manually override this setting
# if you have local administration rights.
# =================================================================================

Connect-IPPSSession -Credential $adminCredentials
# =================================================================================
# Get all Retention Policies in the tenant
# =================================================================================

$retention = get-retentioncompliancepolicy -ExcludeTeamsPolicy

# =================================================================================
# Get the date and perform a Regex replacement of any special chars
# Date and time to millisecond becomes a unique value and is passed in to the temp
# and final filenames in the temp directory
# =================================================================================
$UNIQUED = Get-Date -format "dd-MMM-yyyy HH:mm:ss:ms"
$Unique = $UNIQUED -replace '[^0-9a-zA-Z]+', ''
$xlfile = "$env:TEMP\ComplianceLabelReport"+$Unique+".CSV"
$xlfile2 = "$env:TEMP\ComplianceLabelReport"+$Unique+".xlsx"
$xlfile3 = "$env:TEMP\ComplianceLabelReport"+$Unique+"3.CSV"
$xlfile4 = "$env:TEMP\ComplianceLabelReport"+$Unique+"4.CSV"
Remove-Item $xlfile -ErrorAction SilentlyContinue
new-item $xlfile

# =================================================================================
# Get all SharePoint online sites in the tenant
# =================================================================================
$SPOSites = get-sposite -Limit All

# =================================================================================
# For each retention policy in the tenant get any Compliance tags applied.
# =================================================================================

foreach ($r in $retention) 
    {
        $atagconcat = ""
        $L2concat = ""
        $Output = @()
        $policyname = $r.name

        $policysites = get-retentioncompliancepolicy -identity $r.identity -ExcludeTeamsPolicy -DistributionDetail -RetentionRuleTypes | Select-Object -expandproperty SharepointLocation
        $policysites2 = get-retentioncompliancepolicy -ExcludeTeamsPolicy -DistributionDetail -RetentionRuleTypes

#     $Policysitename = $policysites.name | Out-String

            foreach($s in $Sposites)
                {
#               $sURL = $s.URL
                    if ($policysites.name -eq $S.url)
                        {

# =================================================================================
# Compliance label info can be found either under ComplianceTagProperty or 
# ApplyComplianceTag in the Compliance rule depending on the method of application
# =================================================================================

                            $LabelInfo = Get-RetentionComplianceRule -policy $Policyname | Select-Object ComplianceTagProperty
                            $LabelInfo2 = Get-RetentionComplianceRule -policy $Policyname | Select-Object ApplyComplianceTag


                            if ($LabelInfo2.ApplyComplianceTag)
                                {
                                    $applytag = get-compliancetag -id $LabelInfo2.ApplyComplianceTag -erroraction 'silentlycontinue'
                                    foreach ( $a in $applytag)

# =================================================================================
# ComplianceTag is returned as a path to the tag on the FFO server so the result is 
# split by the character '/' and the end is returned as a variable. This is done for
# each result and then concatenated
# =================================================================================

                                        {
                                            $atag = $a.identity
                                            $atag2 = ($atag -split '/',5)[-1]
                                            $atagconcat = "$atagconcat`n$atag2"
                                            $L2Concat = $atagconcat.trim()
                                        }
                                }

                            foreach ($l in $LabelInfo)
                                {

# =================================================================================
# ComplianceTag is collected using the same split method as before and concatenated
# only if the value returned in this iteration of the loop does not already exist
# within the variable i.e. no duplication.
# =================================================================================

                                    $L1 = $l.ComplianceTagProperty
                                    $l2 = ($l1 -split ',',2)[-1]
                                    $L2Trim = $L2.Trim()
                                    if (($L2Concat -match $L2Trim) -eq "True")
                                        {
                                        }
                                    Else
                                        {
                                            $L2Concat = "$L2concat`n$L2Trim"
                                        }
                                }

# =================================================================================
# IF no labels are applied at all in the policy it is a retention policy, if at
# Least one label is applied it is a label policy
# =================================================================================

                            if(!$LabelInfo.ComplianceTagProperty -and !$Labelinfo2.ApplyComplianceTag)
                                {
                                    $PolicyType = "Retention Policy"
                                }
                            Else
                                {
                                    $PolicyType = "Label Policy"
                                }

# =================================================================================
# PSObject created to house policy information before output to a temp CSV file
# =================================================================================

                            $Output = New-Object -TypeName PSObject
                            $Output | Add-Member -MemberType NoteProperty -Name 'Site_URL' -Value $S.url.Trim()
                            $Output | Add-Member -MemberType NoteProperty -Name 'Policy_Name' -Value $PolicyName
                            $Output | Add-Member -MemberType NoteProperty -Name 'Labels_In_Policy' -Value $L2Concat.trim()
                            $Output | Add-Member -MemberType NoteProperty -Name 'PolicyType' -Value $PolicyType
  
# =================================================================================
# Secondary array created to combine looped iterations of $Output as secondary 
# Variable for CSV output. CSV is later reimported to be exported as Excel file
# =================================================================================
                            $reportoutput = @()
 
                            $Arrayreportoutput = $reportoutput + $Output 
                            import-csv $xlfile -erroraction 'silentlycontinue'
                            If (!$xlfile)
                                {
                                    Write-Output "$Site_URL,Policy_Name,Labels_In_Policy,PolicyType"
		                            $arrayreportoutput| Export-Csv $xlfile -append
	                            }
	                        Else
	                            {
		                            $arrayreportoutput| Export-Csv $xlfile -append
	                            }
                        }
                    $i++
                }
    }


# =================================================================================
# For each policy returned
# =================================================================================

foreach ($policy in $policysites2)
    {
	    $atagconcat = ""
	    $L2concat = ""
	    $policyname = $policy.name

# =================================================================================
# Return the locations that the policy is applied
# =================================================================================

	    $policycheck = get-retentioncompliancepolicy -identity $policyname -ExcludeTeamsPolicy -DistributionDetail -RetentionRuleTypes | Select-Object -expandproperty SharepointLocation

# =================================================================================
# Refine this variable down to only Policies that include "All" SharePoint sites
# =================================================================================

	    $policyALL = $policycheck | Where-Object{$_.DisplayName -eq "All"}
	        if ($policyALL.DisplayName -eq "All")
	            {

# =================================================================================
# Now, get all information on the labels associated with these tenant wide policies
# =================================================================================

                    $LabelInfo = Get-RetentionComplianceRule -policy $Policyname | Select-Object ComplianceTagProperty
                    $LabelInfo2 = Get-RetentionComplianceRule -policy $Policyname | Select-Object ApplyComplianceTag


                    if ($LabelInfo2.ApplyComplianceTag)
                        {
                            $applytag = get-compliancetag -id $LabelInfo2.ApplyComplianceTag -erroraction 'silentlycontinue'
                                foreach ( $a in $applytag)
                                    {
                                        $atag = $a.identity
                                        $atag2 = ($atag -split '/',5)[-1]
                                        $atagconcat = "$atagconcat`n$atag2"
                                        $L2Concat = $atagconcat.trim()
                                    }


                        }

            foreach ($l in $LabelInfo)
                {
                    $L1 = $l.ComplianceTagProperty
                    $l2 = ($l1 -split ',',2)[-1]
                    $L2Trim = $L2.Trim()

                    if (($L2Concat -match $L2Trim) -eq "True")
                        {
                        }
                    Else
                        {
                            $L2Concat = "$L2concat`n$L2Trim"
                        }
                }

            if(!$LabelInfo.ComplianceTagProperty -and !$Labelinfo2.ApplyComplianceTag)
                {
                    $PolicyType = "Retention Policy"
                }
            Else
                {
                    $PolicyType = "Label Policy"
                }

# =================================================================================
# New PSObject created to hold Tenant-Wide policies, appended to $Xlfile  CSV
# =================================================================================

            $Output = New-Object -TypeName PSObject
            $Output | Add-Member -MemberType NoteProperty -Name 'Site_URL' -Value "Tenant-Wide"
            $Output | Add-Member -MemberType NoteProperty -Name 'Policy_Name' -Value $PolicyName
            $Output | Add-Member -MemberType NoteProperty -Name 'Labels_In_Policy' -Value $L2Concat.Trim()
            $Output | Add-Member -MemberType NoteProperty -Name 'PolicyType' -Value $PolicyType
  
 
            $reportoutput = @()
 
            $Arrayreportoutput = $reportoutput + $Output 
            import-csv $xlfile -erroraction 'silentlycontinue'
            If (!$xlfile)
                {
                    Write-Output "$Site_URL,Policy_Name,Labels_In_Policy,PolicyType"
		            $arrayreportoutput| Export-Csv $xlfile -append
	            }
	        Else
	            {
		            $arrayreportoutput| Export-Csv $xlfile -append
	            }
            }
    }

# =================================================================================
# Get all Retention Compliance Policies in the tenant
# =================================================================================

$Rcompliance = Get-Retentioncompliancepolicy -ExcludeTeamsPolicy

$RetentionCompliancePolicies = @()
foreach ($r in $rcompliance)
    {
        $atagconcat = ""
        $L2concat = ""
        $x2concat = ""
        $y2concat = ""
#        $RFL = Get-Retentioncompliancepolicy -identity $r.name -ExcludeTeamsPolicy | FL

# =================================================================================
# Expand properties of Retention compliane policy and bind to variables
# =================================================================================

        $LabelInfo = Get-RetentionComplianceRule -policy $R.name| Select-Object ComplianceTagProperty
        $LabelInfo2 = Get-RetentionComplianceRule -policy $R.name | Select-Object ApplyComplianceTag
        $RSPLoc = get-retentioncompliancepolicy $R.name -DistributionDetail -RetentionRuleTypes -ExcludeTeamsPolicy | Select-Object -expandproperty Sharepointlocation
        $RSPLocEx = get-retentioncompliancepolicy $R.name -DistributionDetail -RetentionRuleTypes -ExcludeTeamsPolicy | Select-Object -expandproperty SharepointlocationException

# =================================================================================
# Get all Labels associated with policy
# =================================================================================

        if ($LabelInfo2.ApplyComplianceTag)
            {
                $applytag = get-compliancetag -id $LabelInfo2.ApplyComplianceTag -erroraction 'silentlycontinue'
                foreach ( $a in $applytag)
                    {
                        $atag = $a.identity
                        $atag2 = ($atag -split '/',5)[-1]
                        $atagconcat = "$atagconcat`n$atag2"
                        $L2Concat = $atagconcat.trim()
                    }
            }

        foreach ($l in $LabelInfo)
            {
                $L1 = $l.ComplianceTagProperty
                $l2 = ($l1 -split ',',2)[-1]
                $L2Trim = $L2.Trim()

                if (($L2Concat -match $L2Trim) -eq "True")
                {
                }
                Else
                {
                    $L2Concat = "$L2concat`n$L2Trim"
                }
                }
            if(!$LabelInfo.ComplianceTagProperty -and !$Labelinfo2.ApplyComplianceTag)
                {
                    $PolicyType = "Retention Policy"
                }
                Else
                {
                    $PolicyType = "Label Policy"
                }

            foreach ($y in $RSPLoc)
                {
                    $y1 = $y.name
                    $y2 = ($y1 -split ',',2)[-1]
                    $y2Trim = $y2.Trim()

                    if (($y2Concat -match $y2Trim) -eq "True")
                        {
                        }
                        Else
                        {
                            $y2Concat = "$y2concat`n$y2Trim"
                        }
                }

            foreach ($x in $RSPLocEX)
                {
                    $x1 = $x.name
                    $x2 = ($x1 -split ',',2)[-1]
                    $x2Trim = $x2.Trim()
                if (($x2Concat -match $x2Trim) -eq "True")
                    {
                    }
                Else
                    {
                        $x2Concat = "$x2concat`n$x2Trim"
                    }
                }

# =================================================================================
# New PS Object to house information on Policies
# =================================================================================

            $RetentionCompliancePolicies = New-Object -TypeName PSObject
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'Name' -Value $R.Name
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'Enabled?' -Value $R.Enabled
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'Labels_In_Policy' -Value $L2Concat.trim()
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'PolicyType' -Value $PolicyType
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'SharePointLocation' -Value $y2Concat.trim()
            $RetentionCompliancePolicies | Add-Member -MemberType NoteProperty -Name 'SharePointLocationException' -Value $x2Concat
 
            $reportoutput2 = @()
 
            $Arrayreportoutput2 = $reportoutput2 + $RetentionCompliancePolicies 
            import-csv $xlfile -erroraction 'silentlycontinue'
            If (!$xlfile)
                {
                    Write-Output "$Site_URL,Policy_Name,Labels_In_Policy,PolicyType"
		            $arrayreportoutput2| Export-Csv $xlfile3 -append
	            }
	        Else
	            {
		            $arrayreportoutput2| Export-Csv $xlfile3 -append
	            }
    }

# =================================================================================
# Get all Retention Labels in the tenant
# =================================================================================

$comptag = Get-Compliancetag

Foreach ($c in $comptag)
    {

        $cidentity = ($c.identity -split '/',5)[-1]

# =================================================================================
# Create PS Object to hold this information
# =================================================================================

        $Compliancetagobj = New-Object -TypeName PSObject
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Name' -Value $cidentity
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Notes' -Value $c.notes
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Retention Action' -Value $c.retentionaction
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Retention Type' -Value $c.Retentiontype
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Regulatory' -Value $c.regulatory
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Retention Duration (Days)' -Value $c.RetentionDuration
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Mode' -Value $c.mode
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Reviewer?' -Value $c.ReviewerEmail
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Created by' -Value $c.Createdby
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Creation Date' -Value $c.Whencreated
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Last modified by?' -Value $c.Lastmodifiedby
        $Compliancetagobj | Add-Member -MemberType NoteProperty -Name 'Modified Date' -Value $c.Whenchanged

        $reportoutput3 = @()
 
        $Arrayreportoutput3 = $reportoutput3 + $Compliancetagobj 
        import-csv $xlfile -erroraction 'silentlycontinue'
        If (!$xlfile)
            {
                Write-Output "$Site_URL,Policy_Name,Labels_In_Policy,PolicyType"
		        $arrayreportoutput3| Export-Csv $xlfile4 -append
	        }
	    Else
	        {
		        $arrayreportoutput3| Export-Csv $xlfile4 -append
            }
    }

# =================================================================================
# Collate 3 exported CSV files in to one xlsx file across 3 worksheets.
# =================================================================================

Import-CSV $xlfile3 | export-excel $xlfile2 -WorkSheetname 'Retention and Label Policies' -StartColumn 1 -Append
Import-CSV $xlfile4 | export-excel $xlfile2 -WorkSheetname 'Label information' -StartColumn 1 -Append
Import-CSV $xlfile | export-excel $xlfile2 -WorkSheetname 'Policies by SharePoint Site' -StartColumn 1 -Append

# =================================================================================
# Open the .XLSX file
# =================================================================================

Invoke-item -path $xlfile2
