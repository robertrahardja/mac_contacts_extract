-- Export all contacts to CSV file
set csvFile to (path to desktop as text) & "contacts_export.csv"
set csvContent to "First Name,Last Name,Organization,Email,Phone" & return

tell application "Contacts"
    set allPeople to people
    
    repeat with aPerson in allPeople
        set rowData to ""
        
        -- First Name
        try
            set rowData to rowData & "\"" & (first name of aPerson) & "\","
        on error
            set rowData to rowData & "\"\","
        end try
        
        -- Last Name
        try
            set rowData to rowData & "\"" & (last name of aPerson) & "\","
        on error
            set rowData to rowData & "\"\","
        end try
        
        -- Organization
        try
            set rowData to rowData & "\"" & (organization of aPerson) & "\","
        on error
            set rowData to rowData & "\"\","
        end try
        
        -- First Email
        try
            set emailList to emails of aPerson
            if (count of emailList) > 0 then
                set rowData to rowData & "\"" & (value of item 1 of emailList) & "\","
            else
                set rowData to rowData & "\"\","
            end if
        on error
            set rowData to rowData & "\"\","
        end try
        
        -- First Phone
        try
            set phoneList to phones of aPerson
            if (count of phoneList) > 0 then
                set rowData to rowData & "\"" & (value of item 1 of phoneList) & "\""
            else
                set rowData to rowData & "\"\""
            end if
        on error
            set rowData to rowData & "\"\""
        end try
        
        set csvContent to csvContent & rowData & return
    end repeat
end tell

-- Write to file
try
    set fileRef to open for access file csvFile with write permission
    set eof of fileRef to 0
    write csvContent to fileRef
    close access fileRef
    
    return "Export complete: " & csvFile
on error
    try
        close access file csvFile
    end try
    return "Export failed"
end try
