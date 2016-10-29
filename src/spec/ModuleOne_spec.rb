# require 'rspec'
require_relative '../../lib/excelLib'

def getSettingsBook
  puts 'getting settings called'
  setup_excel=ExcelLib.new '../dataEngine/000_Setup.xls'
  return setup_excel
end

def getModuleBook
  puts 'getting settings called'
  setup_excel=ExcelLib.new '../dataEngine/Module one.xls'
  return setup_excel
end

def getTestCases testScenario
  testcases=Array.new
  module_excel=getModuleBook
  testcases=module_excel.getTestCases testScenario
  puts "To get list of data from sheet for  #{testScenario}"
  return testcases
end

def getData testScenario
  puts "To get data from sheet for  #{testScenario}"
end

describe 'To test Rspec and spread sheet features' do

  before(:all) do
  end

  before(:each) do
    # puts '2'
    puts 'Getting setup data'
    setup_excel=getSettingsBook
    setupData=setup_excel.getSetupdata 'Setup'
    puts 'Launching ' + setupData['Browser'] + ' : ' + setupData['URL']
  end

  # ======================================= Scenario One =========================================================== #

   describe 'Scenario One' do
       setup_excel=getSettingsBook
       if  setup_excel.scenarioEnabled 'Scenario One'
         testcases=getTestCases 'Scenario One'
         puts testcases

         testcases.each do |testcase|
           it 'test_ruby_rspec with scenario one '+testcase do
             puts 'S1 started with '+testcase
             getData testcase
             puts 'this is a test for '+testcase
             puts 'S1 Ended with '+testcase
           end
         end
       end
   end

  # ======================================= Scenario Two =========================================================== #

  describe 'Scenario Two' do
    setup_excel=getSettingsBook
    if  setup_excel.scenarioEnabled 'Scenario Two'
      testcases=getTestCases 'Scenario Two'
      puts testcases

      testcases.each do |testcase|
        it 'test_ruby_rspec with scenario one '+testcase do
          puts 'S2 started with '+testcase
          getData testcase
          puts 'this is a test for '+testcase
          puts 'S2 Ended with '+testcase
        end
      end
    end
  end


  after(:each) do
    puts 'Closing firefox using after each'
  end

  after(:all) do
    puts 'Doing nothing using after class'
  end

end

