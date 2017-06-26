############################################################################################################
# KEYWORD DRIVEN AUTOMATION TESTING FRAMEWORK IMPLEMENTATION IN RUBY 1.8.7 and WATIR
# 
# AUTHOR : VIJAY SIVASANKARAN ; email:vijaix@gmail.com ; Cell:+1-404-863-1576
# 
# LICENSE: GNU GPLv2
#
# SYSTEM REQUIREMENTS: Windows XP+, Ms Excel 2003+, IE 7+, Ruby 1.8.7 and Watir 1.6.7
# WARNING: NOT SUPPORTED ON RUBY 1.9.x . SET $PATH TO 1.8.7 INCASE YOU HAVE 1.9.x INSTALLED.
#
# IF YOU RECIEVED THIS AS A ZIPFILE,IT SHOULD CONTAIN THE FOLLOWING 4 FILES:
# 1. THIS SOURCE FILE - KEDAFI.rb
# 2. THE EXCEL SPREADSHEET WHICH SERVES AS THE KEYWORD SOURCE/SERVER AND AS RESULTS LOG - ATDriverResults.xls
# 3. RUBY 1.8.7 Installer
# 4. README.txt - A README FILE TO GET YOU STARTED on how to install WATIR and how to input keywords
#    in ATDriverResults.xls and the folder\path to place ATDriverResults.xls
# REMEMBER:ATDriverResults.xls should always be present in C:\KEDAFI.
#
# KEDAFI IS HOSTED ON GITHUB AND CAN ALWAYS BE OBTAINED FROM https://github.com/vijayMKH/KEDAFI
#
#############################################################################################################
# KEDAFI BEGIN
# Using the win32ole rubygems and watir libs
require 'rubygems'
require 'win32ole'
require 'watir'
require 'watir/ie'
require 'watir/screen_capture'

# KEDAFI will be in global ruby obj scope, not declaring modules/classes,so mixin watir with KEDAFI to call/make use of watir methods
include Watir
include Watir::ScreenCapture

# RUBY GLOBAL VARIABLES, setting all to FALSE, set all of the below to TRUE to debug
$VERBOSE=FALSE
$DEBUG=FALSE
$-s=FALSE
$-c=FALSE
$-d=FALSE
$-v=FALSE
$-w=FALSE
$-y=FALSE

# KEDAFI CORE BEGIN
begin
puts "STARTING KEDAFI - Copyright by Vijay Sivasankaran, 2011 - GNU GPLv2 - email: vijaix@gmail.com"
puts "............................................................................................."
sleep(1)
puts "KEDAFI STARTED"

# OPEN KEYWORD DRIVER EXCEL FILE
excel = WIN32OLE::new('excel.Application')
excel.DisplayAlerts = false
workbook = excel.Workbooks.Open('C:\KEDAFI\ATDriverResults.xls')

# READ/WRITE LOOP OVER WORKBOOK AND ALL WORKSHEETS
@test_case_passed=0
@test_case_failed=0
# i= WorkBook Looper j=WorkSheet Looper
for i in 1 .. workbook.Worksheets.Count
  @worksheet = workbook.Worksheets(i)
  @rowcount = @worksheet.UsedRange.Rows.Count

  for @j in 2..@rowcount
    @keyword =@worksheet.Cells(@j, 1).value
    @object_Prop_Name = @worksheet.Cells(@j, 2).value
    @object_Prop_Value = @worksheet.Cells(@j, 3).value
    @expected_Output = @worksheet.Cells(@j, 4).value
    @parm_01 = @worksheet.Cells(@j, 6).value

    case @keyword
      
      when /^Begin/
      @test_step_passed=0
      @test_step_failed=0
      
      when /^OpenURL/
      @Browser=IE.start(@parm_01)
      @Browser.bring_to_front
      @Browser.maximize
    
      when /^SetText/
      @Browser.text_field(:"#{@object_Prop_Name}", @object_Prop_Value).set(@parm_01)
      
# YOU CAN KEEP EXTENDING KEDAFI BY ADDING HTML ELEMENTS AND THE EVENTS HERE WITHIN THIS CASE STRUCTURE
# JUST LIKE THE SetText CASE ABOVE OR THE ClickButton CASE BELOW
# THERE IS ONLY A FINITE UNIVERSE OF HTML ELEMENTS ON WEBPAGES AND THE HTML 5 SPEC
# ONE YOU WRITE HANDLERS FOR ALL OF THOSE AND ALL OF THEIR EVENTS KEDAFI SHOULD COVER THE ENTIRE UNIVERSE
# OF WEB UI TESTS POSSIBLE FROM THE BROWSER
# http://wtr.rubyforge.org/rdoc/1.6.2/classes/Watir/   WILL GIVE YOU ALL THINGS POSSIBLE
    
      when/^ClickButton/
      @Browser.button(:"#{@object_Prop_Name}", @object_Prop_Value).click
    
      when/^ClickDivToggle/
      @Browser.div(:"#{@object_Prop_Name}", @object_Prop_Value).fireEvent("onclick")
            
      when /^CheckText/
      if @object_Prop_Name == "text"
      if @Browser.text.include?@object_Prop_Value
         set_actout("Present")
      else
         set_actout("Absent")
      end
      end
      
      when/^CloseURL/
      @Browser.close
    
      when/^Result/
      set_test_case_result()
      workbook.SaveAs('C:\KEDAFI\ATDriverResults.xls')
      @Browser.close
      
      else
        "Exit"
    end
    
    def set_actout(actout_parmstring)
      if actout_parmstring == "Present"
           @actual_Output="#{@object_Prop_Value}"+" - "+"Present"
           @worksheet.Cells(@j, 5)['Value']=@actual_Output
           set_test_set_result("Pass")
      else
           @actual_Output="#{@object_Prop_Value}"+" - "+"Absent"
           @worksheet.Cells(@j, 5)['Value']=@actual_Output
           set_test_set_result("Fail")
      end
    end
    
    def set_test_set_result(result_parmstring)
      if result_parmstring == "Pass"
           @test_step_passed=@test_step_passed+1
           @test_step_result="Pass"
           @worksheet.Cells(@j,8)['Value']=@test_step_result
           step_result_screen_capture()
      else
           @test_step_failed=@test_step_failed+1
	   @test_step_result="Fail"
           @worksheet.Cells(@j,8)['Value']=@test_step_result
           step_result_screen_capture()
      end
    end
    
    def step_result_screen_capture()
      time_stamp_s = Time.new.strftime('%m%d_%H%M_%S')
      screenshot_filename="C\:\\KEDAFI\\"+"#{time_stamp_s}_#{@actual_Output}"+"\.bmp"
      screen_capture(screenshot_filename,active_window_only=false,save_as_bmp=true)
      @worksheet.Cells(@j, 7)['Value']=screenshot_filename
    end
    
    def set_test_case_result()
      if @test_step_failed > 0
         @test_case_failed=@test_case_failed+1
         @test_case_result="Fail"
         @worksheet.Cells(@j,10)['Value']=@test_case_result
      else
         @test_case_passed=@test_case_passed+1
	 @test_case_result="Pass"
         @worksheet.Cells(@j,10)['Value']=@test_case_result
      end
    end
    
  end #@j
  
end #@i

workbook.Close()
excel.quit()
  
# YOU MIGHT EXTEND THE BELOW METHOD TO REPORT RESULTS TO TFS OR SUBVERSION FOR CONTINOUS INTEGERATION BUILDS
# AS OF NOW THE CURRENT METHOD PRINTS THE TEST SUIT RESULTS TO THE CONSOLE

def put_test_suite_result()
  if @test_case_failed > 0
     puts "Overall Test Suite Result = FAIL"
     puts "Total # of Test Cases = " + (@test_case_passed + @test_case_failed).to_s
     puts "Test Cases Passed = " + @test_case_passed.to_s
     puts "Test Cases Failed = " + @test_case_failed.to_s
  else
     puts "Overall Test Suite Result = PASS"
     puts "Total # of Test Cases = " + (@test_case_passed + @test_case_failed).to_s
     puts "Test Cases Passed = " + @test_case_passed.to_s
  end
end

put_test_suite_result()

# EXCEPTION HANDLER - EXPECT IT TO BE THROWN WHEN YOU HAVE THE DRIVER FILE OPEN AND READ ONLY
rescue
  puts "KEDAFI has TERMINATED ABNORMALLY. Please run program again with $DEBUG set to TRUE to debug."
  workbook.Close()
  excel.quit()
end

# KEDAFI CORE END
# KEDAFI END
