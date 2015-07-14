require 'axlsx'
require 'tempfile'


class ExcelController < ApplicationController

  skip_before_action :verify_authenticity_token
  def download

    @begin
      if params[:extension] == 'xlsx'
        b = Roo::Excelx.new(params[:path])
      else
        b = Roo::Excel.new(params[:path])
      end

      @casos_prueba = Array.new
      @us = Array.new

      b.default_sheet = b.sheets.to_a.at(params[:cu].to_i)

      2.upto(b.last_row) do |line|

        cp_name = b.cell(line,3).to_s
        @casos_prueba.append cp_name
        @us.append b.cell(line,2)

      end

      cu = b.cell(1,1)

    #rescue
    #  redirect_to '/'
    #end
    num_cp = 1

    last_us = @us.at(0)

    short_cu = ''

    0.upto(@casos_prueba.size - 1 ) do |i|
      if @us.at(i).to_i != last_us
        last_us = @us.at(i)
        num_cp = 1
      else
        num_cp = num_cp.to_i + 1;
      end

      if @us.at(i).to_i < 10
        us = '0' + @us.at(i).to_s
      end

      if num_cp.to_i < 10
        num_cp = '0' + num_cp.to_s
      end

      @casos_prueba[i] = short_cu.to_s + ' - ' + us.to_s.split('.')[0] + ' - ' + num_cp.to_s + ' - ' + @casos_prueba.at(i).to_s
    end


    sprint = params[:sprint].to_i

    if sprint < 10
      sprint = '0' + sprint.to_s
    end


    p = Axlsx::Package.new

    # Required for use with numbers
    p.use_shared_strings = true

    p.workbook do |wb|
      # define your regular styles
      styles = wb.styles

      header = styles.add_style :bg_color => '002EB8', :fg_color => 'FF', :b => true, :border => Axlsx::STYLE_THIN_BORDER, :alignment => { :horizontal => :center,
                                                                                                                                           :vertical => :center ,
                                                                                                                                           :wrap_text => false}

      content_1_left = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER,  :alignment => { :horizontal => :left,
                                                                                                                                                   :vertical => :center ,
                                                                                                                                                   :wrap_text => true}
      content_1_no_warp = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER,  :alignment => { :horizontal => :left,
                                                                                                                                                        :vertical => :center ,
                                                                                                                                                        :wrap_text => false}

      content_1_middle = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER,  :alignment => { :horizontal => :center,
                                                                                                                                                   :vertical => :center ,
                                                                                                                                                   :wrap_text => true}

      content_1_fecha = styles.add_style :bg_color => 'B0B0B0', :format_code => 'DD/MM/YYYY', :border => Axlsx::STYLE_THIN_BORDER, :alignment => { :horizontal => :center,
                                                                                                                                                                                     :vertical => :center ,
                                                                                                                                                                                     :wrap_text => true}


      wb.add_worksheet(:name => 'Diseño de CP') do  |ws|
        #Agregamos el encabezado

        con_prioridad = true
        altura_calculada = 120
        fila_impar = [content_1_middle, content_1_middle, content_1_middle, content_1_no_warp, content_1_middle, content_1_middle, content_1_fecha, content_1_middle, content_1_left, content_1_left]
        if params[:prioridad]
          ws.add_row ['CU',	'US',	'Subject',	'Test Name',	'Descripcion',	'Prioridad',	'Fecha',	'Step Name',	'Step Description',	'Expected Result'], :style => header

          0.upto(@casos_prueba.size - 1 ) do |i|
            us = @us.at(i).to_s.split('.')[0]
            if us.to_i < 10
              us = '0' + us.to_s
            end

            subject = '1 - Pruebas Funcionales\Sprint ' + sprint.to_s + '\\' + cu.to_s + '\\' + us.to_s

            ws.add_row [cu,	@us.at(i),	subject,	@casos_prueba.at(i),	'Descripcion',	'3 - Media',	Date.today,	'Step 1',	'Step Description',	'<Expected Result>'], :style => fila_impar, :height=> altura_calculada



          end

          ws.add_data_validation('F2:F' + @us.size.to_s, {
              :type => :list,
              :formula1 =>'"1 - Critico" "2 - Alta" "3 - Media" "4 - Baja"',
              :showDropDown => false})


        else
          ws.add_row ['CU',	'US',	'Subject',	'Test Name',	'Descripcion',	'Fecha',	'Step Name',	'Step Description',	'Expected Result'], :style => header
        end

        ws.column_info.to_a.at(0).width = 24
        ws.column_info.to_a.at(1).width = 5
        ws.column_info.to_a.at(2).width = 17
        # You can merge cells!
        #ws.merge_cells 'A1:C1'

      end

      if params[:prioridad]
        wb.add_worksheet(:name => 'Validaciones - Data') do  |ws|
          ws.add_row ['Prioridad']
          ws.add_row ['1 - Critico']
          ws.add_row ['2 - Alta']
          ws.add_row ['3 - Media']
          ws.add_row ['4 - Baja']
        end
      end

    end

    p.serialize("#{Rails.root}/tmp/excel_out/diseño.xlsx")

    send_file "#{Rails.root}/tmp/excel_out/diseño.xlsx"

  end

  def upload

    file =  params[:analisis]

#file.original_filename

    @extension = file.original_filename.split('.').last

    tmp_file = Tempfile.new([file.original_filename,".#{@extension}"], "#{Rails.root}/tmp/excel_in/")

    File.open(tmp_file, 'wb') do |f|
      f.write file.read
    end

    tmp_file.close

    @casos_usos = Array.new

    if @extension == 'xlsx'
      s = Roo::Excelx.new(tmp_file.path)
    else
      s = Roo::Excel.new(tmp_file.path)
    end

    @path = tmp_file.path

    s.each_with_pagename do |name, sheet|
      #p name
      cu_name = sheet.cell(1,1).to_s
      @casos_usos.append cu_name
    end
    @casos_usos.delete_at @casos_usos.size - 1

  end
end