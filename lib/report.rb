require 'tiny_tds'
require 'activerecord-sqlserver-adapter'

class Newstar < ActiveRecord::Base
  self.primary_key = 'RowID'

  self.pluralize_table_names = false
  self.table_name_prefix = 'dbo.'

  self.establish_connection(
    :adapter    => "sqlserver",
    :host       => "bhmnewstar.signature.local",
    :database   => "informXL_dm",
    :username   => "analyzer",
    :password   => "xxx",
    :timeout    => "60000",
    :persistent => "true"
  )
end

class Report
  require 'date'
  require 'active_support/core_ext/date/calculations'
  require 'active_support/core_ext/numeric/time'
  require 'active_support/core_ext/date/acts_like'
  require 'axlsx'

  def initialize(company,cutoff)
    @company = company
    @cutoff = Date.strptime(cutoff,"%m/%d/%Y")
    @month_beginning = @cutoff.beginning_of_month

    @filename = "#{File.expand_path File.dirname(__FILE__)}/../tmp/AccrualReport.xlsx"
    @p = Axlsx::Package.new
    @wb = @p.workbook
    @summary_table = Hash.new(0)
  end

  def filename
    @filename
  end

  def create
    begin
      tries ||= 5
      entries = Newstar.connection.select_all("SELECT distinct apd.[groupno] as groupno
          , apd.[unitno] as unitno
          , [apheader].[postingdate] as postingdate
          , addy.[name] as name
          , [apheader].[invno] as invno
          , [apheader].[invdate] as invdate
          , [apheader].[amount] as amount
          , apc.[apalloc] as apalloc
       FROM [HBLive].[rems].[apheader] with (nolock)
       left join [HBLive].[rems].[apdetail] as apd
         on [apheader].compcode = apd.compcode and [apheader].voucher = apd.voucher
       left join [HBLive].[rems].[apcat] as apc
         on [apheader].[compcode] = apc.[compcode] and [apheader].[catcode] = apc.[catcode]
       left join [HBLive].[rems].[address] as addy
         on [apheader].[aid] = addy.[aid]
      where [apheader].postingdate >= '#{@month_beginning}' and [apheader].postingdate <= '#{@cutoff}'
        and [apheader].compcode = '#{@company}'
        and [apheader].invdate < '#{@month_beginning}'
        and apd.transno = 0
        and [apheader].updated = 1
      order by apc.[apalloc],apd.groupno, apd.unitno")
    rescue => exception
      if (tries -= 1) > 0
        sleep 90
        retry
      end
    end

    @wb.styles do |s|
      black_cell = s.add_style :bg_color => "00", :fg_color => "FF"
      align_right = s.add_style :alignment => { :horizontal=> :right }
      text_cell = s.add_style :types => [:string]
      currency = s.add_style :num_fmt => 8, :alignment => { :horizontal => :right }

      @wb.add_worksheet(:name => "Detail") do |sheet|
        sheet.add_row ["Group No.", "Unit No.", "Posting Date", "Company","Invoice No.","Invoice Date","Invoice Amount","Trans No.","Allocation","Ref. Allocation","Source Amount"], :style => [black_cell,black_cell,black_cell,black_cell,black_cell,black_cell,black_cell,black_cell,black_cell,black_cell,black_cell]

        entries.each do |e|

          details = get_ap_item_detail(e['groupno'],e['unitno'],e['postingdate']).to_ary
          first = details.shift
          gstring = "G,#{first['compcode']},#{first['divcode']},#{first['acctcode']}"
          gstring << ",#{first['slcode']}" if !first['slcode'].empty?

          @summary_table[gstring] += first['srcamt']

          sheet.add_row ["#{e['groupno']}",
                         "#{e['unitno']}",
                         e['postingdate'].strftime("%m/%d/%Y"),
                         e['name'],
                         e['invno'],
                         e['invdate'].strftime("%m/%d/%Y"),
                         e['amount'],
                         first['transno'],
                         gstring,
                         first['joballoc'],
                         first['srcamt']],
                         :types => [:string,:string,:string,:string,:string,:string,nil,:string,:string,nil,nil],
                         :style => [nil,nil,align_right,nil,nil,align_right,currency,align_right,nil,nil,currency]
      
          details.each do |d|
            gstring = "G,#{d['compcode']},#{d['divcode']},#{d['acctcode']}"
            gstring << ",#{d['slcode']}" if !d['slcode'].empty?

            @summary_table[gstring] += d['srcamt']

            sheet.add_row ["","","","","","","",
                           "#{d['transno']}",
                           "#{gstring}",
                           "#{d['joballoc']}",
                           d['srcamt']],
                           :style => [nil,nil,align_right,nil,nil,align_right,currency,nil,nil,nil,currency]
          end
          sheet.add_row 
        end
      end

      @wb.add_worksheet(:name => "Summary") do |sheet|
        sheet.add_row ["Allocation","Account Name","Source Amount"], :style => [black_cell,black_cell,black_cell]

        sorted_summary_table = @summary_table.to_a.sort! {|a,b| a[0] <=> b[0]}

        total = 0.00
        sorted_summary_table.each do |entry|
          account_name = get_account_name(entry[0])
          total += entry[1].to_f
          sheet.add_row [entry[0],account_name,entry[1]],
               :types => [:string,:string,nil],
               :style => [nil,nil,currency]
        end  
        sheet.add_row
        sheet.add_row [nil,"Total:",total],
             :style => [nil,nil,currency]
      end
    end

    @p.serialize @filename
  end

  private

  def format_big_number(number)
    return number.to_s.reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
  end

  def get_company_name(company)
    name = Newstar.connection.select_all("SELECT [short]
           FROM [HBLive].[rems].[company]
          where compcode = '#{company}'")

    return name[0]['short']
  end

  def get_division_name(company,division)
    name = Newstar.connection.select_all("SELECT [short]
           FROM [HBLive].[rems].[division]
          where compcode = '#{company}'
            and divcode = '#{division}'")

    return name[0]['short']
  end

  def get_glaccount_name(account)
    name = Newstar.connection.select_all("SELECT [AccountNameHB]
           FROM [HBLive].[dbo].[Account_List]
          where [AccountCodeHB] = '#{account}'")

    if name.count == 0
      return ""
    elsif account == "1300"
      return "WIP - Work in Progress"
    else
      return name[0]['AccountNameHB']
    end
  end

  def get_subledger_name(subledger)
    name = Newstar.connection.select_all("SELECT [sldesc]
           FROM [HBLive].[rems].[subledger]
          where slcode = '#{subledger}'")

    return name[0]['sldesc']
  end

  def get_account_name(g_string)
    parts = g_string.split(",")
    compcode = get_company_name(parts[1])
    divcode = get_division_name(parts[1],parts[2])
    glaccount = get_glaccount_name(parts[3])
    if parts[4].nil?
      subledger = ""
    else
      subledger = "#{get_subledger_name(parts[4])}, "
    end

    return "#{subledger}#{glaccount}, #{divcode}, #{compcode}"
  end

  def get_ap_item_detail(groupno,unitno,postingdate)
    items = Newstar.connection.select_all("SELECT [compcode]
        ,[divcode]
        ,[acctcode]
        ,[slcode]
        ,[joballoc]
        ,[srcamt]
        ,[groupno]
        ,[unitno]
        ,[transno]
    FROM [HBLive].[rems].[detail]
    where groupno = '#{groupno}'
    and unitno = '#{unitno}'
    and postingdate = postingdate
    order by transno")

    return items
  end
end
