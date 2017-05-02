# Report-Reports.ps1 - Generates reports for reports
# Copyright (C) 2017. All Rights Reserved. Oregon Department of Transportation.

# Opens a dialog and allows the user to select a directory
function Get-Folder ( $Description, $SelectedFolder, $AllowNew )
{
	Add-Type -AssemblyName System.Windows.Forms
	$FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
	$FolderDialog.Description = $Description
	if ( $AllowNew -eq $false ) {
		$FolderDialog.ShowNewFolderButton = $false
	}
	if ( $SelectedFolder ) {
		$FolderDialog.SelectedPath = $SelectedFolder
	}
	if ( $FolderDialog.ShowDialog() -eq "OK" ) {
		return $FolderDialog.SelectedPath
	} else {
		exit
	}
}

# If the string has a part like "Inital Catalog=XYZ" returns XYZ,
# otherwise returns $connectString
function Convert-ConnectionString ( $connectString ) {
	# If there is "Initial Catalog" somewhere in the connection string,
	# then this is a connection to a database, so pull the database connection
	# out. Otherwise, it's probably a reference to a .xml document, so just
	# use the entire connection string.
	if( $connectString.Contains("Initial Catalog") ) {
		# Split the items in the connection string and search for the
		# "Initial Catalog" property. When we find it, set the $value
		# variable to the value of the Initial Catalog
		$items = $connectString.Split(";")
		foreach ( $i in $items ) {
			if ( $i.Contains("Initial Catalog") ) {
				return $i.Split("=")[1]
			}
		}			
	} else {
		return $connectString
	}
}

$reportDir = ""
$outputDir = ""
$startingLocation = Get-Location

# Set the report directory or prompt the user if not provided
if ( $args[0] ) {
	$reportDir = Resolve-Path $args[0]
} else {
	$reportDir = Get-Folder "Select the reports directory." $null $false
}

# Set the output directory or prompt the user if not provided
if ( $args[1] ) {
	$outputDir = Resolve-Path $args[1]
} else {
	$outputDir = Get-Folder "Select the output directory." $startingLocation $true	
}

$reportDatasetFilename = $(Join-Path $outputDir "reportDatasets.csv")
$reportDataSourceFilename = $(Join-Path $outputDir "reportDataSources.csv")
$projectDatasetFilename = $(Join-Path $outputDir "projectDatasets.csv")
$projectDataSourceFilename = $(Join-Path $outputDir "projectDataSources.csv")

function newReportDataset($localName, $sharedName, $source, $isShared, $isUsed) {
	$newObject = New-Object System.Object
	$newObject | Add-Member -type NoteProperty -name localName -Value $localName;
	$newObject | Add-Member -type NoteProperty -name sharedName -Value $sharedName;
	$newObject | Add-Member -type NoteProperty -name source -Value $source;
	$newObject | Add-Member -type NoteProperty -name isShared -Value $isShared;
	$newObject | Add-Member -type NoteProperty -name isUsed -Value $isUsed;
	return $newObject
}

function newReportDataSource($name, $connection, $isShared, $isUsed) {
	$newObject = New-Object System.Object
	$newObject | Add-Member -type NoteProperty -name name -Value $name
	$newObject | Add-Member -type NoteProperty -name connection -Value $connection
	$newObject | Add-Member -type NoteProperty -name isShared -Value $isShared
	$newObject | Add-Member -type NoteProperty -name isUsed -Value $isUsed
	return $newObject
}

function newProjectDataset($name, $source, $isUsed) {
	$newObject = New-Object System.Object
	$newObject | Add-Member -type NoteProperty -name name -Value $name;
	$newObject | Add-Member -type NoteProperty -name source -Value $source;
	$newObject | Add-Member -type NoteProperty -name isUsed -Value $isUsed;
	return $newObject
}

function newProjectDataSource($name, $connection, $isUsed) {
	$newObject = New-Object System.Object
	$newObject | Add-Member -type NoteProperty -name name -Value $name
	$newObject | Add-Member -type NoteProperty -name connection -Value $connection
	$newObject | Add-Member -type NoteProperty -name isUsed -Value $isUsed
	return $newObject
}

try {
	Set-Location $reportDir
	
	echo "Project`tReport`tDataset`tData Source`tConnection`tIs Shared`tIs Used" | Out-File $reportDatasetFilename
	echo "Project`tReport`tData Source`tSource's Source`tConnection`tIs Shared`tIs Used" | Out-File $reportDataSourceFilename
	echo "Project`tDataset`tData Source`tConnection`tIs Used" | Out-File $projectDatasetFilename
	echo "Project`tData Source`tConnection`tIs Used" | Out-File $projectDataSourceFilename
	
	#  Iterate through every directory (project)
	foreach ( $directory in Get-ChildItem | ?{ $_.PSIsContainer} ) {
		Set-Location $directory.FullName
		$project = $directory.BaseName
		echo $project
		
		[System.Collections.ArrayList] $projectDataSources = @()
		[System.Collections.ArrayList] $projectDatasets = @()
		
		# Get all of the shared data sources in this project
		# Store the name of the data source (the filename without extension)
		# and the database that this data source connects to as key/value pairs
		foreach ( $rds in Get-ChildItem *.rds ) {
			[xml] $xml = Get-Content $rds
			$projectDataSources += newProjectDataSource $rds.BaseName $(Convert-ConnectionString $xml.RptDataSource.ConnectionProperties.ConnectString) $false
		}
		
		# Find all shared datasets and store the name of the dataset and the
		# data source it connects to as a key/value pair
		foreach ( $rsd in Get-ChildItem *.rsd ) {
			[xml] $xml = Get-Content $rsd
			$projectDatasets += newProjectDataSet $rsd.BaseName $xml.SharedDataSet.DataSet.Query.DataSourceReference $false
		}
	
		# Iterate through every report in the project
		foreach ( $rdl in Get-ChildItem *.rdl ) {
			$report = $rdl.BaseName			
			
			echo "-- $report"
			
			[xml] $xml = Get-Content $rdl
			# Extract the namespace of this XML document so we can use it later
			$ns = new-object Xml.XmlNamespaceManager $xml.NameTable
			$ns.AddNamespace("rd", $xml.Report.xmlns )

			[System.Collections.ArrayList] $reportDataSources = @()
			
			# Find all data sources
			foreach ( $dataSource in $xml.Report.DataSources.DataSource ) {				
				if ( $dataSource.DataSourceReference ) {
					$reportDataSources += newReportDataSource $dataSource.name $dataSource.DataSourceReference $true $false 
				} else {
					$reportDataSources += newReportDataSource $dataSource.name $(Convert-ConnectionString $dataSource.ConnectionProperties.ConnectString) $false $false
				}			
			}
			
			[System.Collections.ArrayList] $reportDatasets = @()
			
			# Find all datasets
			# Hidden datasets are parameters with queries. Ignore these.
			foreach ( $usedDataset in $($xml.Report.DataSets.DataSet | ?{ $_.Query.Hidden -ne "true" } ) ) {
				if ( $usedDataset.SharedDataSet ) {
					$reportDatasets += newReportDataset $usedDataset.Name $usedDataset.SharedDataSet.SharedDataSetReference $null $true $false
				} else {
					$reportDatasets += newReportDataset $usedDataset.Name $null $usedDataset.Query.DataSourceName $false $false
				}				
			}
			
			# Mark every actually used dataset as used
			# A used dataset is a dataset that appears in a <DataSetName> tag.
			foreach ( $d in $xml.SelectNodes("//rd:DataSetName", $ns)| Select-Object -Property "#text" ) {
				$dataset = $reportDatasets | ?{ $_.localName -eq $d."#text" }

				# If we didn't find the dataset, then it was probably a report parameter
				# so just ignore it
				if ( $dataset ) {
					$dataset.isUsed = $true
					
					# If this is a shared dataset, then mark the project's dataset as used
					if ( $dataset.isShared -eq $true ) {
						
						# If the dataset starts with a slash, then the dataset reference
						# is remote and not part of this projet. Ignore it in this case
						if ( $dataset.sharedName[0] -ne "/" ) {
						
							$pDataset = $projectDatasets | ?{ $_.name -eq $dataset.sharedName }
							$pDataset.isUsed = $true						
							
							# Mark the project dataset's source as used					
							$pDataSource = $projectDataSources | ?{ $_.name -eq $pDataset.source }

							if ( $pDataSource ) {					
								$pDataSource.isUsed = $true					
							}
						}
					} else {
						# This is an embedded dataset. Mark this report's data source as used
						$rDataSource = $reportDataSources | ?{ $_.name -eq $dataset.source}							
						$rDataSource.isUsed = $true
						
						# If the report data source is shared, then also mark the project's data source as used
						if ( $rDataSource.isShared -eq $true ) {
							$pDataSource = $projectDataSources | ?{ $_.name -eq $rDataSource.connection }						
							
							if ( $pDataSource ) {						
								$pDataSource.isUsed = $true
							}
						}
					}
				}
			}
			
			# Add these items to the reports sheet
			foreach ( $d in $reportDatasets ) {

				$sharedOrEmbedded = $( if ( $d.isShared ) {"shared"} else {"embedded"} )
				$used = $( if ( $d.isUsed ) {"yes"} else {"no"} )
				
				$connection = ""
				$source = $d.source
				$rDataSource = $reportDataSources | ?{ $_.name -eq $d.source}
				
				# If this is a shared dataset, figure out what the project's dataset's source is
				if ( $d.isShared -eq $true  ) {
					$pDataset = $projectDatasets | ?{ $_.name -eq $d.sharedName }
					$source = $pDataset.source
					$pDataSource = $projectDataSources | ?{ $_.name -eq $pDataset.source }
					$connection = $pDataSource.connection
				}
				
				# If the data source is a shared dataset, then find the connection of the
				# project's data source
				elseif ( $rDataSource.isShared -eq $true ) {
					# If the data source starts with a "/", then this references a path.
					if ( $rDataSource.connection[0] -eq "/" ) {
						$connection = $rDataSource.connection
					} else {					
						$pDataSource = $projectDataSources | ?{ $_.name -eq $rDataSource.connection }					
						$connection = $pDataSource.connection
					}
				} else {					
					$connection = $rDataSource.connection
				}
				
				Add-Content $reportDatasetFilename "$project`t$report`t$($d.localName)`t$source`t$connection`t$sharedOrEmbedded`t$used"
			}
			
			# Report data sources
			foreach ( $d in $reportDataSources ) {
				$sharedOrEmbedded = $( if ( $d.isShared ) {"shared"} else {"embedded"} )
				$used = $( if ( $d.isUsed ) {"yes"} else {"no"} )
				$connection = $d.connection
				# If the data source is shared, then figure out what the shared data source's connection is
				if ( $d.isShared ) {
					$pDataSource = $projectDataSources | ?{ $_.name -eq $d.connection }
					$connection = $pDataSource.connection
				}
				
				Add-Content $reportDataSourceFilename "$project`t$report`t$($d.name)`t$($d.connection)`t$connection`t$sharedOrEmbedded`t$used"
			}
		}
		
		# Create the project dataset record
		foreach ( $d in $projectDatasets ) {
			$used = $( if ( $d.isUsed ) {"yes"} else {"no"} )
			$pDataSource = $projectDataSources | ?{ $_.name -eq $d.source }
			Add-Content $projectDatasetFilename "$project`t$($d.name)`t$($d.source)`t$($pDataSource.connection)`t$used"
		}
		
		# Create the project data source record
		foreach ( $d in $projectDataSources ) {
			$used = $( if ( $d.isUsed ) {"yes"} else {"no"} )
			Add-Content $projectDataSourceFilename "$project`t$($d.name)`t$($d.connection)`t$used"
		}		
	}
}

finally {
	Set-Location $startingLocation
}