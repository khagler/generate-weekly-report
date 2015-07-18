--
--	Created by: Ken Hagler
--	Created on: 04/13/14 16:07:25
--
--	Copyright (c) 2014 Ken Hagler
--	All Rights Reserved
--

property reportFolders : {}
property reportDB : ""

-- Handler to select which folders to include in the report. Sets the reportFolders property
-- to the list of folders selected by the user.
on chooseFolders()
	-- Start by getting a list of the names of all top-level folders
	tell application "OmniFocus" to set {folderNames} to {name} of folders of the default document
	choose from list folderNames with prompt "Select folders to report on:" with multiple selections allowed without empty selection allowed
	if the result is not false then
		set reportFolders to the result
	end if
end chooseFolders

-- Handler to determine which day the report is for. We want a report as of 5 PM on Friday,
-- and if today is not Friday, use the most recent Friday.
on getDate()
	set today to current date
	repeat while today's weekday is not Friday
		set today to today - 1 * days
	end repeat
	set today's time to (0 + 17 * hours)
	return today
end getDate

-- Handler to find the names of the currently active projects in the folders included in the report.
on getProjects(folderName)
	tell application "OmniFocus"
		set folderObj to folder folderName of default document
		-- Start by getting all the projects with active status in the specified folder
		set refProjects to a reference to (every project of folderObj whose status is active and singleton action holder is false)
		set {projectList} to {name} of refProjects
		
		-- The folder may have subfolders, and if so we need to check them for projects too
		set subFolderObjs to (folders of folderObj)
		repeat with aSubFolderObj in subFolderObjs
			set refProjects to (a reference to (every project of aSubFolderObj whose status is active and singleton action holder is false))
			set {projectList} to {name} of refProjects
		end repeat
	end tell
	return projectList
end getProjects

-- Handler to find the tasks completed up to daysBack days before the date specified by lastDay.
-- The returned list will include the task name, its project, and the date when it was completed.
on getCompletedTasks(lastDay, daysBack)
	set doneList to {}
	tell application "OmniFocus"
		set refDoneInLastWeek to a reference to (flattened tasks of default document where (completion date ≥ lastDay - daysBack))
		set {lstName, lstContext, lstProject, lstFolder, lstDate} to {name, name of its context, name of its containing project, name of folder of its containing project, completion date} of refDoneInLastWeek
		repeat with iTask from 1 to count of lstName
			set taskText to ""
			if item iTask of lstFolder is in reportFolders then
				set {strName, varProject, varDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstDate}
				if varDate is not missing value then set taskText to taskText & short date string of varDate & " - "
				if varProject is not missing value then set taskText to taskText & " [" & varProject & "] - "
				copy taskText & strName to end of doneList
			end if
		end repeat
	end tell
	return doneList
end getCompletedTasks

-- Handler to generate the text of the report in Markdown format.
on makeReport(reportDate, reportLength)
	set outputText to "Current List of Active Projects" & return & "---" & return & reportDate & return & return as Unicode text
	repeat with aFolder in reportFolders
		repeat with projectName in getProjects(aFolder)
			set outputText to outputText & projectName & return
		end repeat
	end repeat
	set outputText to outputText & return & return & "Completed Tasks" & return & "---" & return & return
	repeat with aTask in getCompletedTasks(reportDate, 7 * days)
		set outputText to outputText & aTask & return
	end repeat
	return outputText
end makeReport

-- Start by seeing if we have some folders to look for projects in. If not, prompt the user to select some.
if (count of reportFolders) is 0 then
	chooseFolders()
end if

-- Generate a report for tasks completed in the past seven days.
set reportText to makeReport(getDate(), 7 * days)

-- Save the report in my DEVONthink database for work
tell application id "DNtp"
	open database reportDB
	create record with {name:"Weekly CM Report.md", type:txt, plain text:reportText} in record "Weekly Reports" of database "Symantec"
end tell
