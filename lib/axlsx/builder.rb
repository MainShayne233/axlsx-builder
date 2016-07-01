module Axlsx
  module Builder
    class Object
      def initialize(attrs={})
        attrs.each { |key, val| send("#{key.to_s}=", val) }
      end
    end

    class Cell < Object
      attr_accessor :text
      attr_accessor :title
      attr_accessor :row
      attr_accessor :col
    end

    class Alignment
      attr_accessor :horizontal, :vertical

      def initialize(horizontal: :left, vertical: :center)
        @horizontal = horizontal
        @vertical = vertical
      end
    end

    class Font < Object
      attr_accessor :name, :size, :color

      def initialize(attrs={})
        @name = 'Liberation Sans'
        @size = 10
        super
      end
    end

    class Format < Object
      attr_accessor :color, :font, :style, :borders, :border_thickness, :alignment

      def initialize(attrs={})
        @font = Font.new
        @alignment = Alignment.new
        super
      end
    end

    class Position
      attr_writer :col
      attr_accessor :row

      def initialize(row: 0, col: 0)
        @row = row
        @col = col
      end

      # Must subtract 1 from col as it is 0 based, not 1 based.
      def col
        @col-1
      end
    end

    class Element < Object
      attr_reader :pos
      attr_accessor :text, :format, :borders, :border_thickness,
                    :merge, :comment

      def initialize(attrs={})
        @pos = Position.new
        @format = Format.new
        super
      end

      def pos=(pos)
        @pos = Position.new(row: pos[0], col: pos[1])
      end

      def row; pos.row; end
      def col; pos.col; end
      def style; format.style; end
      def font; format.font; end
      def color; format.color; end
      def h_align; format.alignment.horizontal; end
      def v_align; format.alignment.vertical; end

      def combined_style
        combined_style = {}
        if self.style
          combined_style.merge!(b: true) if self.style.include? :bold
          combined_style.merge!(i: true) if self.style.include? :italic
          combined_style.merge!(u: true) if self.style.include? :underline
          combined_style.merge!(alignment: {horizontal: :center}) if self.style.include? :center
          combined_style.merge!(sz: 12) if self.style.include? :lg_font
        end
        combined_style.merge!(sz: self.font.size)
        combined_style.merge!(font_name: self.font.name)
        combined_style.merge!(bg_color: self.color) if self.color
        combined_style.merge!(fg_color: self.font.color)
        combined_style.merge!(alignment: {horizontal: h_align, vertical: v_align})
        if self.borders == true
          combined_style.merge! border: {style: self.border_thickness || :medium, color: '00000000'}
        elsif self.borders
          combined_style.merge! border: {style: self.border_thickness || :medium, color: '00000000', edges: self.borders}
        end
        combined_style
      end
    end

    class Title < Element
      attr_accessor :list, :hyperlink

      def initialize attrs
        h_align = :center # Default
        v_align = :center # Default
        attrs.each {|key, val| self.send("#{key.to_s}=", val)}
        super
      end

      def combined_style
        combined_style = super
        combined_style.merge!(b: true, border: {style: :thick, color: '00000000'})
        combined_style.merge!(bg_color: self.color) if self.color
        combined_style
      end
    end


    class Blueprint
      attr_accessor :sheet, :elements, :column_titles, :row_titles, :column_titles_start,
                    :row_titles_start, :row_data, :column_data, :max, :column_title_row_height, :format

      def initialize options = {}
        options[:format] ||= Format.new

        self.elements = if options[:elements]
                          options[:elements].map do |elem|
                            elem.merge!(format: options[:format]) unless elem[:format]
                            Element.new(elem)
                          end
                        else
                          []
                        end
        self.column_titles = options[:column_titles] ? options[:column_titles].map{|elem| Title.new(elem)} : []
        self.row_titles = options[:row_titles] ? options[:row_titles].map{|elem| Title.new(elem)} : []
        self.column_titles_start = options[:column_titles_start] || [1,1]
        self.row_titles_start = options[:row_titles_start] || [1,1]
        self.column_title_row_height = options[:column_title_row_height]
      end
    end
  end





  class Worksheet
    attr_accessor :column_title_indexes, :row_title_indexes, :blueprint, :column_data, :row_data

    def build!
      if self.column_data
        self.blueprint.column_data = self.column_data.map{|elem| Builder::Cell.new(elem)}
      end

      if self.row_data
        self.blueprint.row_data = self.row_data.map{|elem| Builder::Cell.new(elem)}
      end

      set_column_title_indexes

      set_max_row_and_column

      set_row_title_indexes

      set_column_widths

      set_row_heights

      place_elements

      place_column_titles

      set_lists_for_column_titles

      place_row_titles

      set_lists_for_row_titles

      set_column_title_row_height

      place_column_data if self.blueprint.column_data

      place_row_data if self.blueprint.row_data

      move_lists_sheet_to_end

      self
    end

    def set_column_title_indexes
      self.column_title_indexes = {}
      self.blueprint.column_titles.each_with_index {|elem, index| self.column_title_indexes[elem.text] = index+self.blueprint.column_titles_start[1]-1}
    end

    def set_row_title_indexes
      self.row_title_indexes = {}
      self.blueprint.row_titles.each_with_index {|elem, index| self.row_title_indexes[elem.text] = index+self.blueprint.row_titles_start[0]-1}
    end

    def set_column_widths
      col_widths = {}

      element_strings = []

      self.blueprint.elements.each do |element|
        longest_line = element.text.split("\n").max_by{|line| line.length}
        element_strings.push [element.col, longest_line || '']
        (1..element.merge).each {|column| col_widths[column] = 0} if element.merge
      end

      element_strings.each do |col, text|
        col_widths[col] ||= 0
        col_widths[col] = text.length if text.length > col_widths[col]
      end

      self.blueprint.column_titles.each_with_index do |elem, index|
        col_widths[index+self.blueprint.column_titles_start[1]-1] ||= 0
        col_widths[index+self.blueprint.column_titles_start[1]-1] = elem.text.length if elem.text.length > col_widths[index+self.blueprint.column_titles_start[1]-1]
      end

      unless self.blueprint.row_titles.empty?
        col_widths[self.blueprint.row_titles_start[1]-1] ||= 0
        max_row_title = self.blueprint.row_titles.max_by{|elem| elem.text.length}.text.length
        col_widths[self.blueprint.row_titles_start[1]-1] = max_row_title if max_row_title > col_widths[self.blueprint.row_titles_start[1]-1]
      end

      if self.blueprint.column_data
        self.blueprint.column_data.each do |elem|
          title_column = self.column_title_indexes[elem.title]
          col_widths[title_column] ||= 0
          col_widths[title_column] = elem.text.length if elem.text.length > col_widths[title_column]
        end
      end

      if self.blueprint.row_data
        self.blueprint.row_data.each do |elem|
          self.blueprint.row_data.select{|data| data.text == elem.text}.each_with_index do |title_data, index|
            column = self.blueprint.row_titles_start[1] + 1 + index
            col_widths[column] ||= 0
            col_widths[column] = title_data.text.length if title_data.text.length > col_widths[column]
          end
        end
      end

      widths = (0..self.blueprint.max[:col].to_i).map{|col| col_widths[col] ? col_widths[col]+4 : 0}
      self.column_widths *widths
    end

    def set_row_heights
      row_heights = {}

      self.blueprint.elements.each do |elem|
        row_heights[elem.row] ||= 0
        height = elem.text.split("\n").count * 10 + (elem.font.size / 3)
        row_heights[elem.row] = height if height > row_heights[elem.row]
      end

      unless self.blueprint.column_titles.empty?
        max_column_title_height = self.blueprint.column_titles.map{|elem| elem.text.split("\n").size}.max * 10 + 10
        row_heights[self.blueprint.column_titles_start[0]] ||= 0
        row_heights[self.blueprint.column_titles_start[0]] = max_column_title_height if max_column_title_height > row_heights[self.blueprint.column_titles_start[0]]
      end


      (0..self.blueprint.max[:row].to_i).each {|row| self.rows[row-1].height = row_heights[row] || 20}
    end

    def place_elements
      self.blueprint.elements.each do |elem|
        self.name_to_cell("#{column(elem.col)}#{elem.row}").value = elem.text
        style = self.styles.add_style elem.combined_style
        self.name_to_cell("#{column(elem.col)}#{elem.row}").style = style
        self.add_comment ref: "#{column(elem.col)}#{elem.row}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
        self.merge_cells "#{column(elem.col)}#{elem.row}:#{column(elem.col+elem.merge)}#{elem.row}" if elem.merge
      end
    end

    def set_max_row_and_column
      self.blueprint.max = {col: 0, row: 0}
      self.blueprint.max[:row] = self.blueprint.elements.max_by{|elem| elem.row}.row unless self.blueprint.elements.empty?
      self.blueprint.max[:col] = self.blueprint.elements.max_by{|elem| elem.col}.col unless self.blueprint.elements.empty?
      self.blueprint.max[:col] = self.blueprint.column_titles.count if self.blueprint.column_titles and self.blueprint.column_titles.count > self.blueprint.max[:col]
      self.blueprint.max[:row] = self.blueprint.row_titles.count if self.blueprint.row_titles and self.blueprint.row_titles.count > self.blueprint.max[:row]
      self.blueprint.max[:row] += self.blueprint.column_titles_start[0] if self.blueprint.column_titles_start
      self.blueprint.max[:col] += self.blueprint.column_titles_start[1] if self.blueprint.column_titles_start
      self.blueprint.max[:row] += self.blueprint.row_titles_start[0] if self.blueprint.row_titles_start
      self.blueprint.max[:col] += self.blueprint.row_titles_start[1] if self.blueprint.row_titles_start
      self.blueprint.max[:row] += self.blueprint.column_data.size if self.blueprint.column_data
      self.blueprint.max[:col] += self.blueprint.row_data.size if self.blueprint.row_data
      self.blueprint.max[:col] += self.blueprint.column_titles.select{|elem| elem.list}.size + self.blueprint.row_titles.select{|elem| elem.list}.size
      (0..self.blueprint.max[:row]+1).each {self.add_row(Array.new(self.blueprint.max[:col].to_i, nil))}
    end

    def place_column_titles
      self.blueprint.column_titles.each_with_index do |elem, index|
        cell = self.name_to_cell("#{column(index+self.blueprint.column_titles_start[1]-1)}#{self.blueprint.column_titles_start[0]}")
        cell.value = elem.text
        style = self.styles.add_style elem.combined_style
        cell.style = style
        self.add_hyperlink location: elem.hyperlink, ref: cell if elem.hyperlink
        self.add_comment ref: "#{column(index+self.blueprint.column_titles_start[1]-1)}#{self.blueprint.column_titles_start[0]}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
      end
    end

    def set_lists_for_column_titles
      existing_titles_with_lists = self.blueprint.column_titles.select{|elem| elem.list}
      return false if existing_titles_with_lists.empty?
      if lists_sheet = current_lists_sheet
        pre_existing_titles_with_lists = []
        lists_sheet.column_title_indexes.each do |title, column|
          row = 2
          list = []
          while cell = lists_sheet.name_to_cell("#{column column}#{row}") and cell.value
            list << cell.value
            row += 1
          end
          pre_existing_titles_with_lists << Axlsx::Builder::Title.new(text: title, list: list)
        end
        titles_with_lists = pre_existing_titles_with_lists + existing_titles_with_lists
        lists_sheet_index = self.workbook.worksheets.index {|sheet| sheet.name == 'Lists'}
        self.workbook.worksheets.delete_at lists_sheet_index
      else
        titles_with_lists = existing_titles_with_lists
      end
      list_titles = []
      list_data = []
      titles_with_lists.each do |elem|
        list_titles = titles_with_lists.map{|elem| {text: elem.text}}
        elem.list.each {|list_item| list_data << {text: list_item, title: elem.text}}
      end
      list_titles.uniq!
      list_data.uniq!
      blueprint = Axlsx::Builder::Blueprint.new column_titles: list_titles
      lists_sheet = self.workbook.add_worksheet name: 'Lists', blueprint: blueprint
      lists_sheet.column_data= list_data
      existing_titles_with_lists.each do |elem|
        100.times do |row|
          list_column = current_lists_sheet.column_title_indexes[elem.text]
          self.add_data_validation("#{column (self.column_title_indexes[elem.text])}#{self.blueprint.column_titles_start[0]+row+1}", {
              type: :list,
              formula1: "Lists!#{column list_column}2:#{column list_column}#{elem.list.size+1}",
              showDropDown: false,
              showErrorMessage: true,
              errorTitle: '',
              errorStyle: :stop,
              showInputMessage: true})
        end
      end
    end

    def place_row_titles
      self.blueprint.row_titles.each_with_index do |elem, index|
        cell = self.name_to_cell("#{column(self.blueprint.row_titles_start[1]-1)}#{self.blueprint.row_titles_start[0]+index}")
        cell.value = elem.text
        style = self.styles.add_style elem.combined_style
        cell.style = style
        self.add_hyperlink location: elem.hyperlink, ref: cell if elem.hyperlink
        self.add_comment ref: "#{column(self.blueprint.row_titles_start[1]-1)}#{self.blueprint.row_titles_start[0]+index}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
      end
    end

    def set_lists_for_row_titles
      existing_titles_with_lists = self.blueprint.row_titles.select{|elem| elem.list}
      return false if existing_titles_with_lists.empty?
      if lists_sheet = current_lists_sheet
        pre_existing_titles_with_lists = []
        lists_sheet.column_title_indexes.each do |title, column|
          row = 2
          list = []
          while cell = lists_sheet.name_to_cell("#{column column}#{row}") and cell.value
            list << cell.value
            row += 1
          end
          pre_existing_titles_with_lists << Builder::Title.new(text: title, list: list)
        end
        titles_with_lists = pre_existing_titles_with_lists + existing_titles_with_lists
        lists_sheet_index = self.workbook.worksheets.index {|sheet| sheet.name == 'Lists'}
        self.workbook.worksheets.delete_at lists_sheet_index
      else
        titles_with_lists = existing_titles_with_lists
      end
      list_titles = []
      list_data = []
      titles_with_lists.each do |elem|
        list_titles = titles_with_lists.map{|elem| {text: elem.text}}
        elem.list.each {|list_item| list_data << {text: list_item, title: elem.text}}
      end
      list_titles.uniq!
      list_data.uniq!
      blueprint = Axlsx::Builder::Blueprint.new column_titles: list_titles
      lists_sheet = self.workbook.add_worksheet name: 'Lists', blueprint: blueprint
      lists_sheet.data column: list_data
      existing_titles_with_lists.each do |elem|
        100.times do |column|
          list_column = current_lists_sheet.column_title_indexes[elem.text]
          self.add_data_validation("#{column self.blueprint.row_titles_start[1]+column}#{self.row_title_indexes[elem.text]+1}", {
              type: :list,
              formula1: "Lists!#{column list_column}2:#{column list_column}#{elem.list.size+1}",
              showDropDown: false,
              showErrorMessage: true,
              errorTitle: '',
              errorStyle: :stop,
              showInputMessage: true})
        end
      end
    end


    def set_column_title_row_height
      return false if self.blueprint.column_titles.empty?
      if self.blueprint.column_title_row_height
        self.rows[self.blueprint.column_titles_start[0]-1].height = column_title_row_height
      else
        most_lines = self.blueprint.column_titles.max_by{|elem| elem.text.split("\n").size}.text.split("\n").size
        self.rows[self.blueprint.column_titles_start[0]-1].height = most_lines * 10 + 10
      end
    end

    def place_column_data
      self.blueprint.column_data.each do |elem|
        index = self.column_title_indexes[elem.title]
        next unless index
        title_column = column index
        if elem.row
          self.blueprint.name_to_cell("#{title_column}#{elem.row + self.blueprint.column_titles_start[0] + 1}").value = elem.text
        else
          row = self.blueprint.column_titles_start[0] + 1
          while self.name_to_cell("#{title_column}#{row}").value
            row += 1
          end
          self.name_to_cell("#{title_column}#{row}").value = elem.text
        end
      end
    end

    def place_row_data
      self.blueprint.row_data.each do |elem|
        title_row = self.row_title_indexes[elem.title]
        column = self.blueprint.row_titles_start[1]
        while self.name_to_cell("#{column column}#{title_row+1}").value
          column += 1
        end
        self.name_to_cell("#{column column}#{title_row+1}").value = elem.text
      end
    end

    def move_lists_sheet_to_end
      return false if self.name == 'Lists'
      sheets = self.workbook.worksheets
      lists_sheet_index = sheets.index { |sheet| sheet.name == 'Lists' }
      if lists_sheet_index
        lists_sheet = sheets[lists_sheet_index]
        sheets.delete_at lists_sheet_index
        sheets << lists_sheet
        sheets.each_with_index {|sheet, index| sheet.workbook.worksheets[index] = sheet}
      end
    end

    def column index
      column = (index % 26 + 65).chr
      column << (index / 26 + 64).chr if index > 25
      column.reverse
    end

    def current_lists_sheet
      self.workbook.sheet_by_name 'Lists'
    end



    def data column: nil, row: nil
      self.column_data = column
      self.row_data = row
      self.build!
    end

  end

  class Workbook
    def add_worksheet(options={})
      worksheet = Worksheet.new(self, options)
      yield worksheet if block_given?
      worksheet.build! if worksheet.blueprint
    end
  end

end



