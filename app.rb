require 'rubygems'
require 'sinatra'
require 'sinatra/reloader'
require 'logger'
require './lib/report.rb'

::Logger.class_eval { alias :write :'<<' }
access_log = ::File.join(::File.dirname(::File.expand_path(__FILE__)),'log',"#{settings.environment}_access.log")
access_logger = ::Logger.new(access_log)
error_logger = ::File.new(::File.join(::File.dirname(::File.expand_path(__FILE__)),'log',"#{settings.environment}_error.log"),"a+")
error_logger.sync = true

configure do
  use ::Rack::CommonLogger, access_logger
end
   
before {
  env["rack.errors"] =  error_logger
}

get '/' do
	@note = nil
  erb :index
end

post '/' do
  report = Report.new(params[:company],params[:cutoff])
  report.create
  send_file(report.filename, :filename=> report.filename)
  erb :index
end
