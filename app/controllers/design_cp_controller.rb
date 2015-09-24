#!/bin/env ruby
# encoding: utf-8

require 'axlsx'
require 'tempfile'



class DesignCpController < ApplicationController

  skip_before_action :verify_authenticity_token
  def index

  end
  def download

    if  request.get?
      redirect_to design_cp_path
    else
      begin

        if params[:path].to_s.split('.').last == 'xlsx'
          b = Roo::Excelx.new(params[:path])
        else
          b = Roo::Excel.new(params[:path])
        end

        @casos_prueba = Array.new
        @us = Array.new

        @precondiciones = Array.new
        @pasos = Array.new

        b.default_sheet = b.sheets.to_a.at(params[:cu].to_i)


        last_us = ''
        2.upto(b.last_row) do |line|

          cp_name = b.cell(line,3).to_s
          @casos_prueba.append cp_name
          @us.append b.cell(line,2)


          if params[:select_precondicion].to_i > 0
            @precondiciones.append params['precondicion_us_' + @us.last.to_i.to_s]
          else
            @precondiciones.append params[:precondicion]
          end

          if params[:select_pasos].to_i > 0
            @pasos.append params['pasos_us_' + @us.last.to_i.to_s]
          else
            @pasos.append params[:pasos]
          end


        end

        cu = params[:nombre_cu]

        num_cp = 0

        last_us = @us.at(0)

        short_cu = params[:nombre_corto].to_s

        0.upto(@casos_prueba.size - 1 ) do |i|
          if @us.at(i).to_i != last_us.to_i
            last_us = @us.at(i)
            num_cp = 1
          else
            num_cp = num_cp.to_i + 1;
          end

          if @us.at(i).to_i < 10
            us = '0' + @us.at(i).to_s
          else
            us = @us.at(i).to_s
          end

          if num_cp.to_i < 100
            if num_cp.to_i < 10
              num_cp = '00' + num_cp.to_s
            else
              num_cp = '0' + num_cp.to_s
            end
          else
            num_cp = num_cp.to_s
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

          header = styles.add_style :bg_color => '002EB8', :font_name => 'Arial', :sz => 11, :fg_color => 'FF', :b => true, :border => Axlsx::STYLE_THIN_BORDER, :alignment => { :horizontal => :center,
                                                                                                                                                                                 :vertical => :center ,
                                                                                                                                                                                 :wrap_text => false}

          content_1_left = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER, :font_name => 'Calibri', :sz => 8,  :alignment => { :horizontal => :left,
                                                                                                                                                            :vertical => :center ,
                                                                                                                                                            :wrap_text => true}

          content_1_left_par = styles.add_style :bg_color => 'FFFFFF', :border => Axlsx::STYLE_THIN_BORDER, :font_name => 'Calibri', :sz => 8,  :alignment => { :horizontal => :left,
                                                                                                                                                                :vertical => :center ,
                                                                                                                                                                :wrap_text => true}

          content_1_no_warp = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER, :font_name => 'Arial', :sz => 12, :b => true, :alignment => { :horizontal => :left,
                                                                                                                                                                         :vertical => :center ,
                                                                                                                                                                         :wrap_text => false}

          content_1_no_warp_par = styles.add_style :bg_color => 'FFFFFF', :border => Axlsx::STYLE_THIN_BORDER, :font_name => 'Arial', :sz => 12, :b => true,  :alignment => { :horizontal => :left,
                                                                                                                                                                              :vertical => :center ,
                                                                                                                                                                              :wrap_text => false}


          content_1_middle = styles.add_style :bg_color => 'B0B0B0', :border => Axlsx::STYLE_THIN_BORDER,  :alignment => { :horizontal => :center,
                                                                                                                           :vertical => :center ,
                                                                                                                           :wrap_text => true}

          content_1_middle_par = styles.add_style :bg_color => 'FFFFFF', :border => Axlsx::STYLE_THIN_BORDER,  :alignment => { :horizontal => :center,
                                                                                                                               :vertical => :center ,
                                                                                                                               :wrap_text => true}

          content_1_fecha = styles.add_style :bg_color => 'B0B0B0', :format_code => 'DD/MM/YYYY', :border => Axlsx::STYLE_THIN_BORDER, :alignment => { :horizontal => :center,
                                                                                                                                                       :vertical => :center ,
                                                                                                                                                       :wrap_text => true}

          content_1_fecha_par = styles.add_style :bg_color => 'FFFFFF', :format_code => 'DD/MM/YYYY', :border => Axlsx::STYLE_THIN_BORDER, :alignment => { :horizontal => :center,
                                                                                                                                                           :vertical => :center ,
                                                                                                                                                           :wrap_text => true}


          wb.add_worksheet(:name => 'Diseño de CP') do  |ws|
            #Agregamos el encabezado

            con_prioridad = true
            altura_calculada = 120
            fila_impar = [content_1_middle, content_1_middle, content_1_left, content_1_no_warp, content_1_left, content_1_middle, content_1_fecha, content_1_middle, content_1_left, content_1_left]
            fila_par = [content_1_middle_par, content_1_middle_par, content_1_left_par, content_1_no_warp_par, content_1_left_par, content_1_middle_par, content_1_fecha_par, content_1_middle_par, content_1_left_par, content_1_left_par]

            fila_impar_sin_prioridad = [content_1_middle, content_1_middle, content_1_left, content_1_no_warp, content_1_left, content_1_fecha, content_1_middle, content_1_left, content_1_left]

            fila_par_sin_prioridad = [content_1_middle_par, content_1_middle_par, content_1_left_par, content_1_no_warp_par, content_1_left_par, content_1_fecha_par, content_1_middle_par, content_1_left_par, content_1_left_par]
            if params[:prioridad]
              ws.add_row ['CU',	'US',	'Subject',	'Test Name',	'Descripcion',	'Prioridad',	'Fecha',	'Step Name',	'Step Description',	'Expected Result'], :style => header

              0.upto(@casos_prueba.size - 1 ) do |i|
                us = @us.at(i).to_s.split('.')[0]
                if us.to_i < 10
                  us = '0' + us.to_s
                end
                descripcion = "Precondición: \n" + "#{@precondiciones.at(i).to_s}"
                subject = '1 - Pruebas Funcionales\Sprint ' + sprint.to_s + '\\' + cu.to_s + '\\' + us.to_s

                if i.odd?

                  ws.add_row [cu,	@us.at(i),	subject,	@casos_prueba.at(i),	descripcion,	'3 - Media',	Date.today,	'Step 1',	@pasos.at(i).to_s,	''], :style => fila_impar, :height=> altura_calculada

                else

                  ws.add_row [cu,	@us.at(i),	subject,	@casos_prueba.at(i),	descripcion,	'3 - Media',	Date.today,	'Step 1',	@pasos.at(i).to_s,	''], :style => fila_par, :height=> altura_calculada

                end

              end

              ws.add_data_validation('F2:F' + @us.size.to_s, {
                  :type => :list,
                  :formula1 =>"='Validaciones - Data'!$A$2:$A$5",
                  :showDropDown => false})


            else
              ws.add_row ['CU',	'US',	'Subject',	'Test Name',	'Descripcion',	'Fecha',	'Step Name',	'Step Description',	'Expected Result'], :style => header

              0.upto(@casos_prueba.size - 1 ) do |i|
                us = @us.at(i).to_s.split('.')[0]
                if us.to_i < 10
                  us = '0' + us.to_s
                end
                descripcion = "\n\nPrecondición: \n" + "#{@precondiciones.at(i).to_s}" + "\n\nResultado Esperado: \n \n"
                subject = '1 - Pruebas Funcionales\Sprint ' + sprint.to_s + '\\' + cu.to_s + '\\' + us.to_s

                if  i.odd?

                  ws.add_row [cu,	@us.at(i),	subject,	@casos_prueba.at(i),	descripcion,	Date.today,	'Step 1', @pasos.at(i).to_s,	''], :style => fila_impar_sin_prioridad, :height=> altura_calculada

                else

                  ws.add_row [cu,	@us.at(i),	subject,	@casos_prueba.at(i),	descripcion,	Date.today,	'Step 1',	@pasos.at(i).to_s,	''], :style => fila_par_sin_prioridad, :height=> altura_calculada

                end


              end


            end

            ws.column_info.to_a.at(0).width = 24
            ws.column_info.to_a.at(1).width = 5
            ws.column_info.to_a.at(2).width = 17
            ws.column_info.to_a.at(4).width = 30

            if params[:prioridad]
              ws.column_info.to_a.at(8).width = 30
            else
              ws.column_info.to_a.at(7).width = 30
            end

            # You can merge cells!
            #ws.merge_cells 'A1:C1'

          end

          if params[:prioridad]
            wb.add_worksheet(:name => 'Validaciones - Data') do  |ws|
              ws.add_row ['Prioridad'] , :style => header
              ws.add_row ['1 - Critico']
              ws.add_row ['2 - Alta']
              ws.add_row ['3 - Media']
              ws.add_row ['4 - Baja']
            end
          end

        end

        tmpfile = Tempfile.new(['diseño','.xlsx'], "#{Rails.root}/tmp/")

        p.serialize tmpfile.path

        send_file tmpfile.path, :filename => short_cu.to_s + '.xlsx'

      rescue
        redirect_to design_cp_path, :alert => 'El archvivo fue removido por favor vuelva a subirlo'
      end
    end

  end

  def upload

    if  request.get?
      redirect_to '/'
    else

      file =  params[:analisis]

      @extension = file.original_filename.to_s.split('.').last

      path = "#{Rails.root}/tmp/"
      filename = "#{file.original_filename.split('.').first}#{Process.pid}.#{@extension}"
      tmp_file = Rails.root.join(path, filename)

      File.open(tmp_file, 'wb') do |f|
        f.write file.read
        f.close
      end



      @casos_usos = Array.new
      if @extension == 'xlsx'
        s = Roo::Excelx.new(tmp_file.to_s)
      else
        s = Roo::Excel.new(tmp_file.to_s)
      end
      @path = tmp_file

      s.each_with_pagename do |name, sheet|
        #p name
        cu_name = sheet.cell(1,1).to_s
        @casos_usos.append cu_name
      end
      @casos_usos.delete_at @casos_usos.size - 1

    end
  end


  def upload_confirmation

    if  request.get?
      redirect_to '/'
    else

      @path = params[:path]

      @extension = params[:extension].to_s

      @prioridad = params[:prioridad]

      if params[:path].to_s.split('.').last == 'xlsx'
        b = Roo::Excelx.new(params[:path])
      else
        b = Roo::Excel.new(params[:path])
      end

      b.default_sheet = b.sheets.to_a.at(params[:cu].to_i)


      @cu = params[:cu]
      @us = Array.new

      2.upto(b.last_row) do |line|

        @us.append b.cell(line,2)

      end

      @us.uniq!

      @cu_name = params[:nombre_cu].to_s

      @cu_name_short = params[:nombre_corto].to_s

      @sprint = params[:sprint]

    end
  end
end
