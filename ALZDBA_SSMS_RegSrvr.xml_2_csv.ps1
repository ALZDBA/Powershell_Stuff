<#
Extract data from RegSrvr.xml to generate CSV files for upload to SSMS Registered servers using "SSMS - CreateSqlServerRegistrationsInSSMS.ps1" 

https://sqlbenjamin.wordpress.com/2018/08/07/creating-registered-servers-in-ssms-via-powershell/

If your profile got messed up, and you haven't been able to export your SSMS registered servers, but still have the RegSrvr.xml available, you can recover your registrations using this workaround.

This is a quick and dirty solution ( as we only use ADAuthentication for DBAs ) to process "C:\Users\<yourAccount>\AppData\Roaming\Microsoft\SQL Server Management Studio\RegSrvr.xml"

#>
set-location C:\SSMS_Regsrvr
Clear-Host


[xml]$SSMSRegSRVR = get-content -Path '.\RegSrvr.xml'

#Please don't mind the way to primitive usage of XML in this script. You can see it just isn't my dada to consume XML as it should.

$AllRegSRVR = $SSMSRegSRVR.model.bufferSchema.definitions.document.data.Schema.bufferData.instances.document

$RegSRVR = foreach ( $reg in $AllRegSRVR.data ) {
    #$reg.RegisteredServersStore.ServerGroups.collection.Reference 
    $RS = $reg.RegisteredServer
    if ( $RS ) {
        $ParentURI = $rs.Parent.Reference.Uri
        # Parent,ServerName,DatabaseName,DisplayName,ConnectionString
        1 | Select @{n='Parent';e={$ParentURI.replace('/RegisteredServersStore/ServerGroup/DatabaseEngineServerGroup/ServerGroup/','')}}, @{n='DisplayName';e={$RS.Name.'#text'}},@{n='Description';e={$RS.Description.'#text'}},@{n='ServerName';e={$RS.ServerName.'#text'}} , @{n='DatabaseName';e={'master'}},@{n='ServerType';e={$RS.ServerType.'#text'}},@{n='AuthenticationType';e={$RS.AuthenticationType.'#text'}}
        }
    }


$TargetFile = $('{0}\RegSRVR_{1}.csv' -f (get-location).Path , (get-date -Format 'yyyyMMdd_HHmmss'));

#Export All in one file
#$RegSRVR | sort Parent, DisplayName | export-csv -Path $TargetFile -NoClobber -NoTypeInformation -Delimiter ',' ;


# Split results to file per SSMS registered servrs branch
$TargetFolder = $('{0}\MijnRegs_{1}' -f (get-location).Path , (get-date -Format 'yyyyMMdd_HHmmss'));
md $TargetFolder
$List = @();
$TargetFileTemplate = $('{0}\GroupName.csv' -f $TargetFolder );
$ref = '';
foreach ( $s in $RegSRVR ) {
    if ( $ref -eq $s.Parent ) {
        $List += $s 
        }
    else {
        if ( $ref -ne '' ) {
            if ( $List.count -gt 0 ) {
                $TargetFileGroup = $TargetFileTemplate.Replace('GroupName',$ref);
                $List | Select ServerName,DatabaseName,DisplayName | sort DisplayName | export-csv -Path $TargetFileGroup -NoClobber -NoTypeInformation -Delimiter ',' ;

                }
            $List = @();
            }
        $ref = $s.Parent
        $List += $s 
         
        }
    }

if ( $List.count -gt 0 ) {
    $TargetFileGroup = $TargetFileTemplate.Replace('GroupName',$ref);
    $List | sort Parent, Name | export-csv -Path $TargetFileGroup -NoClobber -NoTypeInformation -Delimiter ',' ;

    }


<#
Import result CSV into SSMS registered servers ( thank you Benjamin Reynolds )

& .\CreateSqlServerRegistrationsInSSMS.ps1 -Path $TargetFolder -HasHeaderRow -OverwriteGroup

#>
