require 'spreadsheet'

class ExcelLib

  def initialize filepath
    # puts "This is to check Excel lib initialize method with #{filepath}"
      # book = Spreadsheet.open '../dataEngine/000_Setup.xls'
    @book = Spreadsheet.open filepath
    # puts "check done with #{filepath}"
  end

  def getSetupdata setUpSheetname
    @setupsheet = @book.worksheet setUpSheetname
    setupData=Hash.new
    @setupsheet.each do |row|
      setupData.store(row[0],row[1])
    end
     return setupData
  end

  def scenarioEnabled scenarioName
     # puts 'Hello from ScenarioEnabled'
     runMode=false
     @scenariossheet = @book.worksheet 'Test Scenarios'
     scenarioRow = getRowContains scenarioName,getcellIndex('Test Scenario ID','Test Scenarios'),'Test Scenarios'
     # puts scenarioRow
     row=@scenariossheet.row scenarioRow
    # puts '1r1'
    #  puts row[getcellIndex 'Runmode','Test Scenarios']
     if row[getcellIndex 'Runmode','Test Scenarios'] == 'Yes'
       runMode=true
     end
    # puts '2r2'
     return runMode
  end

  def getcellIndex coloumnName,sheetname

    ColumnDictionary sheetname
    cellIndex = @dict[coloumnName]
    #puts cellIndex
    return cellIndex

  end

  def ColumnDictionary sheetName
    @scenariossheet = @book.worksheet sheetName
    @dict=Hash.new
    # puts 'printing row size1'
     for col in 0...@scenariossheet.row(0).size
      @dict.store(getCelldata(0,col,sheetName),col)
     end
    # puts @dict
    # puts 'printing row size2'
  end

  def getRowContains testCaseName,colNum,sheetName
    #  puts 'from getRowContains',testCaseName,colNum,sheetName
      @scenariossheet1 = @book.worksheet sheetName
      for rowIndex in 1...@scenariossheet1.row_count
        # puts rowIndex
        # puts getCelldata(rowIndex,colNum,sheetName)
         if getCelldata(rowIndex,colNum,sheetName) == testCaseName
         #  puts 'Test case found at',rowIndex
           break
          end
      end
    return rowIndex
  end

  def getTestCases testscenario
    testcasesList = Array.new
    startCaserow=getRowContains(testscenario, getcellIndex('Test Scenario ID','Test Cases'),'Test Cases')
    # puts 'Print start test row : ',startCaserow
    lastCaserow=getTestcaseCount('Test Cases',testscenario,startCaserow)
    # puts 'printing last row here : ',lastCaserow
    for row in startCaserow...lastCaserow
      if getCelldata(row,getcellIndex('Runmode','Test Cases'),'Test Cases')=='Yes'
        testcasesList<< getCelldata(row,getcellIndex('Test Case','Test Cases'),'Test Cases')
      end
    end
    # puts testcasesList
    return testcasesList

  end

  def getTestcaseCount sheetName,testscenID,testCaseStart
    testCasesheet = @book.worksheet sheetName
    for row in testCaseStart...testCasesheet.row_count
      # puts 'Printing scenario id we are searching for : ',testscenID
      if testscenID!=getCelldata(row,0,sheetName)
        lastrow=row
        # puts 'To test last row return 1'
         return lastrow
      end
    end
    # puts 'To test last row return 2'
    return testCasesheet.row_count
  end

  def getTestdata testcase,sheetName
    dataset= Hash.new
    testCasesheet = @book.worksheet sheetName
    row=getRowContains(testcase,getcellIndex("Test Case",sheetName),sheetName)
    for col in 1...testCasesheet.row(row).size
      dataset.store(getCelldata(0,col,sheetName),getCelldata(row,col,sheetName))
    end
    return dataset
  end

  def getCelldata rownum,collnum,sheetName
   # puts 'getCelldata',rownum,collnum,sheetName
    @scenariossheet = @book.worksheet sheetName
    cellValue = @scenariossheet.row(rownum)[collnum]
    return cellValue
  end

end