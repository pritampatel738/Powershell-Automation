### Handle the case when data is coming from multiple datasets ###
### Connecting to Azure database ###
### Getting data from power bi workspace ###

# set execution policy so that one can run this script ....
Set-ExecutionPolicy Bypass -Scope Process -Force

# install the required software to connect power bi to powershell ....
if (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt) {} 
else {
    Write-Host 'Required Module does not exist, installing Required Module MicrosoftPowerBIMgmt'
    if(Get-Module -ListAvailable -Name NuGet){
    }
    else{
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    }
    Install-Module -Name MicrosoftPowerBIMgmt -Force
}

# connect to power bi account which can access the required workspace .....
Connect-PowerBIServiceAccount

### Change the database if you want to connect to any other database ###
$name1 = "test" # this is database name under which pbiData and pbiDataOrg tables are present
$workspaceID = Read-Host -Prompt "Input your WorkspaceID"


################ Calling rest-api method to get refreshed date ###################
# $url = "https://api.powerbi.com/v1.0/myorg/datasets/396632be-385a-4aa9-89b3-f3d832600a57/refreshes/?`$top=1"

###### If you want details of any group then below url will work fine. #######
# #$url = "https://api.powerbi.com/v1.0/myorg/groups/me/datasets/8c1c9515-94ed-454c-8fd3-9826743b3581/refreshes"

# $refreshes = Invoke-PowerBIRestMethod -Url $url  -Method GET
# # echo $refreshes
# # echo $refreshes.value
# $orgVal = $refreshes | ConvertFrom-Json
# #echo $orgVal.value
# echo $orgVal.value.endtime




# getting powerbi imports  ...... get-powerbiimport -WorkspaceId 
# getting powerbi datasets ...... get-powerbidataset 
$dataset = get-powerbidataset -WorkspaceId $WorkspaceID #f4c97034-a420-4678-9b81-634673bd5b68
#$imports = get-powerbiimport #-WorkspaceId 9ffdcded-787b-473b-9ff1-1556f55a5624
#$datasource = get-powerbidatasource -datasetid 7af82c13-4c47-4dcd-a2d6-ceff585cd5eb
#echo $imports

#echo $dataset
echo "`nHere is the list of all datasets on this workspace `n"

foreach($var in $dataset){
    echo $var.Name
}
echo "`n `n"

#echo "Here is your all the import information `n"
#foreach($var in $imports){
	#echo $var.id 
	#echo $var.Name   #$var.Updated.DateTime 
	#echo $var.Reports.Name
	#echo $var.Datasets.Name
	#echo $var.Created.DateTime
	#echo $var.Updated.DateTime
	#echo `n
#}

### Set the location to sql server (Optional in some cases) ###
#set-location "sqlserver:\sql\MAQ22304\sqlexpress"


#### Truncate the temporary(pbiData) table made on Azure. #############
$params0 = @{
	 'Database' = $name1

	'ServerInstance' =  'ocpinsightsdevsql.database.windows.net'

	'Username' = 'ocpadmindev'

	'Password' = 'OCP!nsights@123'

	#'OutputSqlErrors' = $true
	 'query' = "
				 truncate table pbidata
				 GO
				 "
	 }
 invoke-sqlcmd @params0
 echo "Truncation Done"
 ######### Truncation Done. #############
 
 
 ## Create a function to return the days difference between two dates ...
 # function DateDiff{
	# # subtract $date1 from $date2 ...
	# param ([String] $date1,[String] $date2)
	
	# $count = 0;
	
	# $day1 = [int]$date1.substring(8,2)
	# $day2 = [int]$date2.substring(8,2)
	
	# $month1 = [int]$date1.substring(5,2)
	# $month2 = [int]$date2.substring(5,2)
	
	# $year1 = [int]$date1.substring(0,4)
	# $year2 = [int]$date2.substring(0,4)
	
	# $yearDiff = $year2 - $year1
	# $monthDiff = $month2 - $month1
	# $dayDiff = $day2 - $day1 
	
	# # calculate the approximate date ....
	# $count = 365 * $yearDiff + 30 * $monthDiff + ($day2 - $day1)
	
	
	# return $count
 
 # }



########################### Upload the data to Azure. ####################################


#foreach($var in $imports){

	# # if dataset name matches in $dataset and isRefreshable is true then it'a an import and keep it...
	#$datasetName = $var.Datasets.Name # name of the dataset 
	#echo $datasetName
	#foreach($datasetName in $var.Datasets.Name){
		foreach($var1 in $dataset){  # check for all the dataset and take only import models ....
			#echo $var1.IsRefreshable $var1.Name
			#$count = 0
			#$var1.Name -eq $datasetName -AND 
			if($var1.IsRefreshable -eq 'True'){
				#echo $datasetName
				#$name = [String]$var.Datasets.Name
				$name = $var1.Name
				echo "Getting data for $($name):"
				#	echo $var.Datasets.Name
				$id = [String]$var1.Id  # id of the dataset ....
				
				
				# $url = "https://api.powerbi.com/v1.0/myorg/datasets/$($id)/refreshes"
				# $refreshes = Invoke-PowerBIRestMethod -Url $url  -Method GET
				
				# $orgVal = $refreshes | ConvertFrom-Json
				#echo $orgVal
				
				# $count1 = 0
				# $todayDate = (get-date).ToString("yyyy:MM:dd") # -Format g
				# #$todayDate = $todayDate.ToString().substring(0,10)
				# # format the date ....
				# #$todayDate = $todayDate -format "yyyyMMdd"
				
				# #echo $todayDate
				
				# foreach($var in $orgVal.value.endtime){
					# $pdate = [datetime]$var
					# # format the date ....
					# $pdate = $pdate.ToString("yyyy:MM:dd")
					# #echo $pdate
					# #echo "The date difference is : "
					# $count1 = DateDiff -date1 $pdate -date2 $todayDate
					# #echo $count1
					# if($count1 -le 30){
						# $count = $count + 1
					# }
				# }
				# #echo "The frequency is "
				# #echo $count
				
				
				# $count = [int]$count
				
				
				
				#$url = "https://api.powerbi.com/v1.0/myorg/datasets/$($id)/refreshes" #?`$top=1
				$url = "https://api.powerbi.com/v1.0/myorg/groups/$($WorkspaceID)/datasets/$($id)/refreshes"
				# get the latest refresh time using invoke-powerbirestmethod ....
				#echo "Fetching refresh time from Workspace"
				$refreshes = Invoke-PowerBIRestMethod -Url $url  -Method GET
				#echo $refreshes
				# echo $refreshes.value
				$orgVal = $refreshes | ConvertFrom-Json
				#echo $orgVal.value
				# echo $orgVal
				$date = "";
				foreach($val in $orgVal.value){
					
					
					if($val.status -eq 'Completed'){
						#echo $val.status
						$date = $val.endTime
						break
					}
				 }
				
				
				
				#$date = $orgVal.value.endtime
				#echo "The date is : "
				echo "$($id) `t $($name) `t $($date) `t 1"
				echo ""
				echo ""
				#	echo  $var.id
				# $date = $var.Updated.Date  # last refreshed date ...
				# #	echo $var.Updated.DateTime
				#	$getDate = (Get-Date).ToString()
				
				# ##### For connecting to azure database pass some more parameters as mentioned below #####
				  
					 echo "Uploading data to SQL Database. "
					 $params = @{
					   'Database' = $name1

					   'ServerInstance' =  'ocpinsightsdevsql.database.windows.net'

					   'Username' = 'ocpadmindev'

	                   'Password' = 'OCP!nsights@123'

					  #'OutputSqlErrors' = $true
					   'query' = "
								 insert into [dbo].[pbidata]
								 ([id],[Name],[RefreshedDate],[RefreshFrequency])
								 values 
								 ('$id','$name','$date','1')
								 GO
								 UPDATE pbidata SET ReportingSnapshotDateKey = Getdate()
								 GO
							 "
					 }
					 invoke-sqlcmd @params
					
				}
			#}
	}
#}

############## Upload to Azure Done. ############################



###################### Merge and Update the database on Azure. ######################
$params1 = @{
	'Database' = $name1

	'ServerInstance' =  'ocpinsightsdevsql.database.windows.net'

	'Username' = 'ocpadmindev'

	'Password' = 'OCP!nsights@123'

	#'OutputSqlErrors' = $true
	'query' = 'exec usp_mergepbidata
				GO'
	}
invoke-sqlcmd @params1
echo "Merging Done via USP"

#################### Data Merging and Updation Done. ###########################



######### Collect the data from Azure to be sent to Microsoft Teams ########################
$params1 = @{
	'Database' = $name1

	'ServerInstance' =  'ocpinsightsdevsql.database.windows.net'

	'Username' = 'ocpadmindev'

	'Password' = 'OCP!nsights@123'
 
	#'OutputSqlErrors' = $true 'Please Refresh dataset with id : '+
	'query' = "
                SELECT id AS ID, name AS Name 
				FROM 
				pbidataorg 
				WHERE  datediff(day,refresheddate,reportingsnapshotdatekey) >= RefreshFrequency
				GO"
	}
 $sqlReturn = invoke-sqlcmd @params1
 #echo $sqlReturn.GetType();
 $retVal = ""
 if($sqlReturn.id){
		  echo ""	 
		  echo "Data to be sent to the Teams"
		  foreach($var in $sqlReturn){
                $varID = $var.id.ToString()
                $varName = $var.Name.ToString()
                #echo "$varID `t $varName"
                $retVal = $retVal + $varID + "`t" + $varName + "`n"
            }
	 }
echo $retVal
################ Data has been collected #################



###################### Send the notifications to Microsoft Teams ####################

 $uri = "https://outlook.office.com/webhook/ff84dff0-b7a1-4756-83a2-e087ffb6560e@e4d98dd2-9199-42e5-ba8b-da3e763ede2e/IncomingWebhook/005e6f799c6246deac6b0ee929c44faf/5ba00831-e4ad-457f-ba57-4a0e16d2e8bf"

 $body = ConvertTo-JSON @{
     text = $retVal
 }

 #Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json'

################### Notifications Sent ############################




# # set the location back to where your powershell script resides ....
# #      "C:\Users\prita\Desktop\Pritam_2304\PBITables\pbiInfo.ps1"
# set-location "C:\Users\prita\Desktop\Pritam_2304\PBITables"

#echo "Here is your information about all the dataset that are being used `n"
#foreach($var in $dataset){
#	echo $var
#}


#### connecting to azure database .... ####
	  #'Database' = 'myazuredatabase'

	  #'ServerInstance' =  'yoursqinstance.database.windows.net'

	  #'Username' = 'adam'

	  #'Password' = 'mysecretpassword'

	  #'OutputSqlErrors' = $true


#### Connect to SQL server on local SSMS ####
# set a location
#set-location "sqlserver:\sql\MAQ22304\sqlexpress"

#### Truncate the table and reinsert all the values .....
#$params0 = @{
#	'database' = $name1
#	'serverinstance' = 'MAQ22304\sqlexpress'
#	'query' = "
#				truncate table pbidata
#				GO
#				"
#	}
#invoke-sqlcmd @params0


#foreach($var in $imports){
#	$name = [String]$var.Datasets.Name
#	echo $var.Datasets.Name
#	$id = [String]$var.id
#	echo  $var.id
#	$date = $var.Updated.Date
#	echo $var.Updated.DateTime
#	$getDate = (Get-Date).ToString()
#	$params = @{
#	'database' = $name1
#	'serverinstance' = 'MAQ22304\sqlexpress'
#	'query' = "
#				insert into [dbo].[pbidata]
#				([id],[Name],[RefreshedDate])
#				values 
#				('$id','$name','$date')
#				GO
#				UPDATE pbidata SET currentDate = Getdate()
#				GO
				
#			"
#	}
#	invoke-sqlcmd @params
#}

#$params1 = @{
#	'database' = $name1
#	'serverinstance' = 'MAQ22304\sqlexpress'
#	'query' = 'exec usp_mergepbidata
#				GO'
#	}
#invoke-sqlcmd @params1

#set-location "C:\Users\prita\Desktop\Pritam_2304\PBITables"



#$reports = get-powerbireport
#$reports



































































