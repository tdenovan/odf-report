require 'rubygems'
require 'zip'
require 'fileutils'
require 'nokogiri'

require File.expand_path('../odf-report/parser/default',  __FILE__)
require File.expand_path('../odf-report/parser/markup_parser',  __FILE__)

require File.expand_path('../odf-report/image',                 __FILE__)
require File.expand_path('../odf-report/field',                 __FILE__)
require File.expand_path('../odf-report/text',                  __FILE__)
require File.expand_path('../odf-report/text_field',            __FILE__)
require File.expand_path('../odf-report/file',                  __FILE__)
require File.expand_path('../odf-report/nested',                __FILE__)
require File.expand_path('../odf-report/section',               __FILE__)
require File.expand_path('../odf-report/table',                 __FILE__)
require File.expand_path('../odf-report/chart',                 __FILE__)
require File.expand_path('../odf-report/spreadsheet',           __FILE__)
require File.expand_path('../odf-report/relationship_manager',  __FILE__)
require File.expand_path('../odf-report/image_manager',         __FILE__)
require File.expand_path('../odf-report/slide_manager',         __FILE__)
require File.expand_path('../odf-report/chart_manager',         __FILE__)
require File.expand_path('../odf-report/table_manager',         __FILE__)
require File.expand_path('../odf-report/report',                __FILE__)
require File.expand_path('../odf-report/slide',                __FILE__)
