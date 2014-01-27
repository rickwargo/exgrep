class ExcelObject
  class AbstractExcelObject
    def initialize(workbook)
      # new instance of an excel object attached to the workbook
      @workbook = workbook
      @sheet = ""
      @row = 0
      @col = 0
      @line = 0
    end

    def each_sheet
      @workbook.Worksheets.each do |sht|
        @sheet = sht.Name
        yield sht if @sheet.match(Options.sheets_matching)
      end
    end
    
    def look_in
      raise NotImplementedError
    end
    
    def find_first_row(sht)
      1  # sht.Cells.Find("*", sht.Range("IV65536"), look_in, 2, 1, 1).Row 
    end
    
    def find_first_column(sht)
      1 # sht.Cells.Find("*", sht.Range("IV65536"), look_in, 2, 2, 1).Column 
    end
    
    def find_last_row(sht)
      50 # sht.Cells.Find("*", sht.Range("A1"), look_in, 2, 1, 2).Row 
    end
    
    def find_last_column(sht)
      cell = sht.Cells.Find("*", sht.Range("A1"), look_in, 2, 2, 2)
      cell == nil ? 1 : cell.Column
    end
    
    def used_range(sht)
      first_row = find_first_row(sht)
      first_column = find_first_column(sht)
              
      last_row = find_last_row(sht)
      last_column = find_last_column(sht)
      
      sht.Range(sht.Cells(first_row, first_column), sht.Cells(last_row, last_column)) 
    end
    
    def each_cell
      @workbook.Worksheets.each do |sht|
        @sheet = sht.Name
        if @sheet.match(Options.sheets_matching)
          rng = used_range(sht)
          $stderr.puts "Scanning sheet #{@sheet}!#{rng.Address}" if Options.verbose
          rng.Rows.each do |row|
            @row = row.Row
            row.Columns.each do |col|
              @col = col.Column
              yield col
            end
          end
        end
      end
    end
    
    def each
      raise NotImplementedError
    end
    
    def where_r1c1
      "#{@sheet}!R#{@row}C#{@col}"
    end
    
    def where_a1
      col = @col > 26 ? ((@col - 1)/26 + 64).chr + (((@col - 1) % 26)+25).chr : (@col+64).chr
      "#{@sheet}!#{col}#{@row}"
    end
    
    def where_line
      "#{@line}"
    end
  end

  class CellExcelObject < AbstractExcelObject
    def where
      where_a1
    end
  end
  
  class FindCellExcelObject < CellExcelObject
    def each_cell
      @workbook.Worksheets.each do |sht|
        @sheet = sht.Name
        if @sheet.match(Options.sheets_matching)
          rng = sht.UsedRange
          $stderr.puts "Scanning sheet #{@sheet}!#{rng.Address}" if Options.verbose
					cell = sht.UsedRange.Find("*", nil, look_in, xlPart=2, xlByColumns=2, xlNext=1)
					if !cell.nil?
						first_address = cell.Address
						begin
							@row = cell.Row
							@col = cell.Column
							yield cell
							cell = sht.UsedRange.FindNext(cell)
						end until cell.nil? || cell.Address == first_address
					end
        end
      end
    end
  end
  
  class ValueExcelObject < FindCellExcelObject
    def look_in
      -4163 # xlValues
    end

    def each
      each_cell do |cell|
        val = cell.Value.to_s
        yield val if val != ""
      end
    end
  end

  class CommentsExcelObject < FindCellExcelObject
    def look_in
      -4144 # xlComments
		end

    def each
      each_cell do |cell|
        val = cell.Comment.Text.to_s
        yield val if val != ""
      end
    end
  end

  class FormulaExcelObject < FindCellExcelObject
    def look_in
      -4123 # xlFormulas
    end

    def each
      each_cell do |cell|
        @cell = cell
        yield cell.Formula if cell.HasFormula
      end
    end
    
    def replace(new_formula)
      @cell.Formula = new_formula
    end
  end

  class ConditionalFormattingExcelObject < CellExcelObject
    def used_range(sht)
      sht.UsedRange
    end
    
    def each
      each_cell do |cell|
        cell.FormatConditions.each do |fc|
          begin
            op = fc.Operator
          rescue
            op = 0
          end
          begin
            op1 = fc.Formula1
          rescue
            op1 = ""
          end
          begin
            op2 = fc.Formula2
          rescue
            op2 = ""
          end
          formula = case op
            when 0
              op1
            when 1
              "between(#{op1}, #{op2})"
            when 2
              "not(between(#{op1}, #{op2}))"
            when 3
              "#{op1} == #{op2}"
            when 4
              "#{op1} != #{op2}"
            when 5
              "> #{op1}"
            when 6
              "< #{op1}"
            when 7
              ">= #{op1}"
            when 8
              "<= #{op1}"
          end
          yield formula
        end
      end
    end
  end
  
  class ControlExcelObject < AbstractExcelObject
    def where
      "#{@sheet}.#{@name}"
    end
    
    def each
      each_sheet do |sht|
        sht.OLEObjects.each do |obj|
          obj.ole_methods.each do |prop|
            Options.controls.each do |c|
              @name = "#{obj.Name}.#{prop}" 
              if @name =~ /#{c}/
                @object = obj
                @property = prop.to_s
								val = nil
								begin
									val = @object[@property]
								rescue
									val = nil
								end
                yield val unless val.nil?
              end
            end
          end
        end
      end
    end

    def replace(new_value)
      @object[@property] = new_value
    end
  end

  class LineExcelObject < AbstractExcelObject
    def where
      where_line
    end

    def iterate(objects, &block)
      @line = 0
      objects.each do |obj|
        @line += 1
        begin
          val = obj.Value.to_s
        rescue # in case COM fails us
          val = "?"
        end
        yield obj.Name + ": " + val
      end
    end
  end

  class NameExcelObject < LineExcelObject
    def each(&block)
      iterate @workbook.Names, &block
    end
		
	def replace(line)
		@workbook.Names.Item(@line).Value = line.sub(/^(\w+!)?\w+: =/, '=')
	end
  end

  class PropertyExcelObject < LineExcelObject
    def each(&block)
      iterate @workbook.BuiltinDocumentProperties, &block
    end
  end

  class AddInExcelObject < LineExcelObject
    def each
      @line = 0
      @workbook.Application.AddIns.each do |addin|
        @line += 1
        yield "#{addin.FullName} (installed: #{addin.Installed})"
      end
      @workbook.Application.COMAddIns.each do |addin|
        @line += 1
        yield "COM: #{addin.progID} (#{addin.Description})"
      end
    end
  end

  class ReferencesExcelObject < LineExcelObject
    def each
      @line = 0
      @workbook.Application.VBE.ActiveVBProject.References.each do |ref|
        @line += 1
        yield "#{ref.Name} -- #{ref.Description rescue "<no description>"} (path: #{ref.FullPath})"
      end
    end
  end

  class WorkbookCommentExcelObject < LineExcelObject
    def each(&block)
      @line = 0
      @workbook.Comments.each do |comment|
        @line += 1
        yield comment
      end
      
      each_sheet do |sht|
        iterate sht.Comments, &block
      end
    end
  end

  class MacroExcelObject < LineExcelObject
    def where_line
      "#{@comp}!##{@line}"
    end
    
    def replace(line)
      comp = @workbook.VBProject.VBComponents(@comp)
      comp.CodeModule.ReplaceLine(@line, line)
    end

    def delete_current_line
      comp = @workbook.VBProject.VBComponents(@comp)
      comp.CodeModule.DeleteLines(@line, 1)
    end

    def each
      @workbook.VBProject.VBComponents.each do |comp|
        @comp = comp.Name
        code = comp.CodeModule.Lines(1, 65536)
        @line = 0
        code.split(/\r?\n/).each do |line|
          @line += 1
          yield line
        end
      end
    end
  end

  class ProcExcelObject < MacroExcelObject
    attr_accessor :proc
    
    def where_line
      "#{@comp}.#{@proc}!##{@line}"
    end

    def each
      @workbook.VBProject.VBComponents.each do |comp|
        @comp = comp.Name
        begin
          @line = comp.CodeModule.ProcBodyLine(@proc, 0)
          cnt = comp.CodeModule.ProcCountLines(@proc, 0)
          code = comp.CodeModule.Lines(@line, cnt)
          code.split(/\r?\n/).each do |line|
            yield line
            @line += 1
          end
        rescue
          @line = 0
        end
      end
    end
  end

 class << self
    def new(object_type, workbook)
      klass = 
        case object_type.downcase
          when /^macro/
            MacroExcelObject
          when /^proc/
            ProcExcelObject
          when /^formula/
            FormulaExcelObject
          when /^value/
            ValueExcelObject
          when /^name/
            NameExcelObject
          when /^conditional/
            ConditionalFormattingExcelObject
          when /^propert/
            PropertyExcelObject
          when /^(wb|workbook)comment/
            WorkbookCommentExcelObject
          when /^comment/
            CommentsExcelObject
          when /^addin/
            AddInExcelObject
          when /^reference/
	    ReferencesExcelObject
          when /^control/
            ControlExcelObject
        end
      klass::new(workbook)
    end
  end
end