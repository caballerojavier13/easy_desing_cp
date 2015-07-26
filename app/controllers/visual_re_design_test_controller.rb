#!/bin/env ruby
# encoding: utf-8

require 'axlsx'
require 'tempfile'

class VisualReDesignTestController < ApplicationController
  def index
  end

  def download
    if  request.get?
      redirect_to re_design_test_path
    else
      begin

        file =  params[:design]

        extension = file.original_filename.to_s.split('.').last

        path = "#{Rails.root}/tmp/"

        tmp_file = Tempfile.new([file.original_filename.split('.').first, ".#{extension}"], path)

        File.open(tmp_file, 'wb') do |f|
          f.write file.read
          f.close
        end

        if extension == 'xlsx'
          b = Roo::Excelx.new(tmp_file.path)
        else
          b = Roo::Excel.new(tmp_file.path)
        end

        tmp_file.unlink

        head = Array.new

        offset_1 = 0
        offset_2 = 0
        offset_3 = 0
        offset_4 = 0

        if params[:cu]
          head.append 'CU'
          offset_1 = 1
        end
        if params[:us]
          head.append 'US'
          offset_2 = offset_1  + 1
        end

        head.append 'Subject'
        head.append 'Test Name'
        head.append 'Descripcion'

        if params[:prioridad]
          head.append 'Prioridad'
          offset_3 = offset_2  + 1
        end
        if params[:fecha]
          head.append 'Fecha'
          offset_4 = offset_3  + 1
        end

        head.append 'Step Name'
        head.append 'Step Description'
        head.append	'Expected Result'


        casos_prueba = Array.new

        cp = Array.new

        step_name = Array.new

        2.upto(b.last_row) do |line|

          if params[:cu]
            cp.append b.cell(line, 1).to_s
          end
          if params[:us]
            up.append b.cell(line, 1 + offset_1).to_s
          end

          cp.append b.cell(line,1 + offset_2).to_s
          cp.append b.cell(line,2 + offset_2).to_s
          cp.append b.cell(line,3 + offset_2).to_s

          if params[:prioridad]
            cp.append b.cell(line,4 + offset_2).to_s
          end
          if params[:fecha]
            cp.append b.cell(line,4 + offset_3).to_s
          end

          cp.append b.cell(line,4 + offset_4).to_s
          step_name.append b.cell(line,4 + offset_4).to_s

          cp.append b.cell(line,5 + offset_4).to_s
          cp.append b.cell(line,6 + offset_4).to_s

          casos_prueba.append cp

          cp = nil
          cp = Array.new
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

          estilo_fila_par = Array.new
          estilo_fila_impar = Array.new

          if params[:cu]
            estilo_fila_par.append content_1_middle_par
            estilo_fila_impar.append content_1_middle
          end
          if params[:us]
            estilo_fila_par.append content_1_middle_par
            estilo_fila_impar.append content_1_middle
          end

          estilo_fila_par.append content_1_left_par
          estilo_fila_impar.append content_1_left
          estilo_fila_par.append content_1_no_warp_par
          estilo_fila_impar.append content_1_no_warp
          estilo_fila_par.append content_1_left_par
          estilo_fila_impar.append content_1_left

          if params[:prioridad]
            estilo_fila_par.append content_1_middle_par
            estilo_fila_impar.append content_1_middle
          end
          if params[:fecha]
            estilo_fila_par.append content_1_fecha_par
            estilo_fila_impar.append content_1_fecha
          end

          estilo_fila_par.append content_1_middle_par
          estilo_fila_impar.append content_1_middle
          estilo_fila_par.append content_1_left_par
          estilo_fila_impar.append content_1_left
          estilo_fila_par.append content_1_left_par
          estilo_fila_impar.append content_1_left


          wb.add_worksheet(:name => 'Diseño de CP') do  |ws|
            #Agregamos el encabezado
            altura_calculada = 120


            ws.add_row head, :style => header
            num_cp = 0

            0.upto(casos_prueba.size - 1 ) do |i|

              num_cp += 1 if step_name.at(i).to_s.eql?('Step 1')

              if num_cp.odd?

                ws.add_row casos_prueba.at(i), :style => estilo_fila_impar, :height=> altura_calculada

              else

                ws.add_row casos_prueba.at(i), :style => estilo_fila_par, :height=> altura_calculada

              end

            end


            if params[:prioridad]
              case offset_2
                when 0
                  rango = 'D2:D' + casos_prueba.size.to_s
                when 1
                  rango = 'E2:E' + casos_prueba.size.to_s
                when 2
                  rango = 'F2:F' + casos_prueba.size.to_s
              end

              ws.add_data_validation(rango, {
                                              :type => :list,
                                              :formula1 =>'1 - Critico;2 - Alta;3 - Media;4 - Baja',
                                              :showDropDown => false})
            end

            if params[:cu]
              ws.column_info.to_a.at(0).width = 24
            end
            if params[:us]
              ws.column_info.to_a.at(0 + offset_1).width = 5
            end

            ws.column_info.to_a.at(0 + offset_2).width = 17
            ws.column_info.to_a.at(2 + offset_2).width = 30

            if params[:prioridad]
              ws.column_info.to_a.at(3 + offset_2).width = 20
            end

            ws.column_info.to_a.at(4 + offset_4).width = 30
            ws.column_info.to_a.at(5 + offset_4).width = 30

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

        tmpfile = Tempfile.new(['re_diseño','.xlsx'], "#{Rails.root}/tmp/")

        p.serialize tmpfile.path

        send_file tmpfile.path, :filename => file.original_filename.to_s.split('.').first + ' - Rediseñado.xlsx'

      rescue
        redirect_to re_design_test_path, :alert => 'Ha ocurrido un error con el archivo. Por favor reintente nuevamente.'
      end
    end
  end
end
