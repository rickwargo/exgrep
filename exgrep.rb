#!/usr/bin/ruby
## require 'rubyscript2exe'
require 'win32ole'
require File.dirname(__FILE__) + '/excel_object'
require File.dirname(__FILE__) + '/option_parser'

## require 'excel_object'
## exit if RUBYSCRIPT2EXE.is_compiling?

class Excel 
	@@excel = nil
	
	def initialize
		startup
	end
	
	def application
		@@excel
	end
	
	def startup
		@@excel = @@excel.nil? ? WIN32OLE.new("Excel.Application") : @@excel
		@@excel.Visible = false
		@@excel.EnableEvents = false
		@@excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
		@@excel.DisplayAlerts = false		
	end
	
	def finish
		@@excel.Quit
		sleep 0.25
		@@excel = nil
	end
	
	def recycle
		finish
		startup
	end
end

def grep1(excel, file)
  matches = 0

  if Options.recurse && FileTest.directory?(file)
    grep(excel, Dir["#{file.gsub(/\\/, "/")}/*.xl*"])
  end
		
  if FileTest.file?(file) && !FileTest.directory?(file) && file.match(Options.include) && (Options.exclude.empty? || !file.match(Options.exclude))
	workbook = excel.Workbooks.Open((file =~ /[\\\/]/ ? '' : Dir.getwd + '/') + file)
	$stderr.puts "Problem opening #{file}!" if workbook.nil?
	begin
		needs_save = 0
		Options.search.each do |object_type|
			$stderr.puts "Searching #{object_type} in #{file} for \"#{Options.expression.join("\"|\"")}\"" if Options.verbose
			excel_obj = ExcelObject.new(object_type, workbook)
			excel_obj.proc = Options.procedure if object_type =~ /^proc/
			excel_obj.each do |obj|
				Options.expression.each do |expr|
				opts = 0
				opts |= Regexp::IGNORECASE if Options.ignore_case
				opts |= Regexp::EXTENDED if Options.extended
				opts |= Regexp::MULTILINE if Options.multi_line
				re = Regexp.new(expr, opts)
					if re.match(obj) || Options.invert_match
						if Options.files_with_matches 
							puts file
							return
						elsif Options.delete_matching_line
							$stderr.puts "Deleting line matching #{re} in #{file}" if Options.verbose
							excel_obj.delete_current_line
							needs_save = needs_save + 1
						elsif !Options.replace.empty?
							t = obj.sub(re, Options.replace)
							if t != obj
								$stderr.puts "Replacing #{re} in #{file} with \"#{Options.replace}\"" if Options.verbose
								excel_obj.replace(t)
								needs_save = needs_save + 1
							end
						else
							printf("%s%s%s\n",
								Options.recurse || ARGV.length > 1 ? "[#{file}] " : "",
								Options.line_numbers ? "#{excel_obj.where}: " : "",
								obj)
                
							matches += 1
							return if Options.max_count.to_i > 0 && matches >= Options.max_count.to_i
						end
					end
				end
			end
		end if !workbook.nil?
	ensure
		workbook.Save if !workbook.nil? && needs_save > 0
		workbook.Close(false) if !workbook.nil?
		workbook = nil
	end
	if Options.files_without_matches && matches == 0
		puts file
	end
  end  
end

def grep(xcel, files)
  n = 0
  files.each do |file|
    grep1(xcel.application, file)
	xcel.recycle if Options.recycle_every > 0 && (n+=1) % Options.recycle_every == 0
  end
end

begin # if __FILE__ == $0
  Options = OptionParser.parse(ARGV)
  
	xcel = Excel.new
	begin
		grep(xcel, ARGV)
	ensure
		xcel.finish
	end
end