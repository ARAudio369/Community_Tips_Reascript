-- Set the path to your spreadsheet file
local spreadsheetPath = "PASTE YOUR FILE PATH HERE - KEEP QUOTATIONS"

function readSpreadsheet(filename)
    local file = io.open(filename, "r") 
    if file then
        local data = {}
        for line in file:lines() do
            local cells = {} 
            local cell = ""
            local inQuotes = false 
            for i = 1, #line do
                local char = line:sub(i, i)
                if char == "\"" then
                    inQuotes = not inQuotes 
                elseif char == "," and not inQuotes then
                    table.insert(cells, cell) 
                    cell = "" 
                else
                    cell = cell .. char 
                end
            end
            table.insert(cells, cell) 
            table.insert(data, cells) 
        end
        file:close() 
        return data
    end
    return nil
end

function clearConsole()
    reaper.ShowConsoleMsg("") 
end

function getRandomRow()
    return math.random(1, 50)
end

function displayCells(data, row)
    local rowCells = data[row]
    if rowCells then
        local output = table.concat(rowCells, "\n\n")
        output = output:match("^%s*(.-)%s*$") 
        reaper.ShowConsoleMsg(output .. "\n")
    else
        reaper.ShowConsoleMsg("Selected row does not exist.\n")
    end
end

local function showMessageBox(message, questionMark)
    local buttonFlags = 1
    if questionMark then
        buttonFlags = reaper.ShowMessageBox(message, " ", 3)
    else
        buttonFlags = reaper.ShowMessageBox(message, " ", 1)
    end
    if buttonFlags == 6 then
        return "yes"
    elseif buttonFlags == 7 then
        return "no"
    elseif buttonFlags == 1 then
        return "ok"
    else
        return "cancel"
    end
end

function shouldCloseMessageBox(lineNum, columnNum)
    if lineNum == 6666 and columnNum == 2 then
        return true
    elseif lineNum == 6735 and columnNum == 2 then
        return true
    elseif lineNum == 6734 and columnNum == 3 then
        return true
    else
        return false
    end
end

local data = readSpreadsheet(spreadsheetPath)
if data then
    
    clearConsole()

    local row = getRandomRow()

    displayCells(data, row)

    local executionCount = tonumber(reaper.GetExtState("Script", "ExecutionCount")) or 0
    executionCount = executionCount + 1

    if executionCount >= 50 then

        local file = io.open(spreadsheetPath, "r")
        if file then
            local lines = {}
            for line in file:lines() do
                local row = {}
                for value in line:gmatch("[^,]+") do
                    table.insert(row, value)
                end
                table.insert(lines, row)
            end
            file:close()

            local lineNum = 6666
            local columnNum = 1

            while lineNum <= #lines do
                local row = lines[lineNum]
                if row then
                    local currentCellValue = row[columnNum]
                    if currentCellValue and currentCellValue ~= "" then
                        local questionMark = currentCellValue:sub(-1) == "?"
                        local result = showMessageBox(currentCellValue, questionMark)
                        if result == "yes" then
                            lineNum = lineNum + 1 
                        elseif result == "no" then
                            columnNum = columnNum + 1 
                        elseif result == "ok" then
                           
                            if shouldCloseMessageBox(lineNum, columnNum) then
                                break
                            else
                                if lineNum == 6682 and columnNum == 1 then
                                    lineNum = 6681 
                                    columnNum = 2 
                                elseif lineNum == 6687 and columnNum == 2 then
                                    lineNum = 6683 
                                    columnNum = 3 
                                elseif lineNum == 6691 and columnNum == 3 then
                                    lineNum = 6688 
                                    columnNum = 4 
                                elseif lineNum == 6692 and columnNum == 4 then
                                    lineNum = 6696 
                                    columnNum = 1 
                                elseif lineNum == 6704 and columnNum == 1 then
                                    lineNum = 6708 
                                    columnNum = 2 
                                elseif lineNum == 6708 and columnNum == 1 then
                                    lineNum = 6704 
                                    columnNum = 2 
                                else
                                    lineNum = lineNum + 1 
                                end
                            end
                        elseif result == "cancel" then
                            break
                        end
                    else
                        reaper.ShowMessageBox("Empty cell detected.", "ReaScript", 0)
                        lineNum = lineNum + 1 
                    end
                end
            end

            
            clearConsole()
            reaper.SetExtState("Script", "ExecutionCount", "0", true)
        else
            reaper.ShowMessageBox("Failed to open the spreadsheet file.", "ReaScript", 0)
        end
    else
    
        reaper.SetExtState("Script", "ExecutionCount", tostring(executionCount), true)
    end
else
    reaper.ShowConsoleMsg("Failed to open the spreadsheet file.\n")
end

