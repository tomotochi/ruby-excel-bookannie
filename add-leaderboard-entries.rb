# coding: UTF-8

require "open-uri"
require "nokogiri"
require "win32ole"

#------------------------------------------------------------------------------
# config
#------------------------------------------------------------------------------

now = Time.now
THIS_YEAR = now.strftime("%Y")
TODAY = now.strftime("%Y/%m/%d")

$conf = {
  :db => {
    :filename => "book-annie-#{THIS_YEAR}.xlsx",
    :template => "book-annie-template.xlsx",
  },
  :uri => {
    :daily => proc{|page| "http://www.romancebookcafe.jp/book/novel/ranking/100/class/daily/page/#{page}"},
    :weekly => proc{|page| "http://www.romancebookcafe.jp/book/novel/ranking/100/class/weekly/page/#{page}"},
    :monthly => proc{|page| "http://www.romancebookcafe.jp/book/novel/ranking/100/class/monthly/page/#{page}"},
  },
  :encoding => "UTF-8",
}

#------------------------------------------------------------------------------
# functions
#------------------------------------------------------------------------------

def get_pages(uri, encoding)
  doc = Nokogiri::HTML.parse(open(uri), nil, encoding)
  
  pages = doc.xpath('//p[@class="pagenation"]').inner_text
  pages = pages.scan(/\d+/).collect{|s| s.to_i}.uniq.sort
  puts :pages => pages
  return pages
end

def add_entries(uri, encoding, worksheet)
  doc = Nokogiri::HTML.parse(open(uri), nil, encoding)
  
  # get the last cell
  # http://msdn.microsoft.com/en-us/library/office/aa139976%28v=office.10%29.aspx
  lastrow = worksheet.Cells.SpecialCells(ExcelConst::XlCellTypeLastCell).Row
  range = worksheet.Cells.Range("A#{lastrow + 1}")
  
  # add entries
  i = 1
  doc.xpath('//div[@id="main_contents"]/div[@class="ranking_area"]').each do |node|
    book = {
      :rank => node.xpath('dl/dt[@class="hidden"]').text.scan(/\d+/).first.to_i,
      :title => node.xpath('dl/dd[@class="title"]/a').text,
      :author => node.xpath('dl/dd[@class="title"]/span[@class="author"]/a').text,
      
      :lead => node.xpath('dl/dd[@class="leed"]').text,
      :img => node.xpath('dl/dd[@class="book"]/a/img/@src').text,
    }
    
    puts "#{book[:rank]}\t#{book[:title]}"
    
    row = range.Rows[i]
    row.Columns[1] = TODAY
    row.Columns[2] = book[:title]
    row.Columns[3] = book[:author]
    row.Columns[4] = book[:rank]
    
    i = i.succ
  end
end

#------------------------------------------------------------------------------
# initialization
#------------------------------------------------------------------------------

# create OLE instances
excel = WIN32OLE.new("Excel.Application")
fso = WIN32OLE.new("Scripting.FileSystemObject")

# consts
class ExcelConst; end
WIN32OLE.const_load(excel, ExcelConst)

#------------------------------------------------------------------------------
# main
#------------------------------------------------------------------------------

# ensure file exists
filename = fso.GetAbsolutePathName($conf[:db][:filename])
unless fso.FileExists(filename)
  template = fso.GetAbsolutePathName($conf[:db][:template])
  fso.CopyFile(template, filename)
end

# working with excel
begin
  workbook = excel.Workbooks.Open(:FileName => filename)
  
  # retrieve the leaderboard's html and concat entries
  
  # get daily leaderboard
  pages = get_pages($conf[:uri][:daily][1], $conf[:encoding])
  pages.each do |page|
     add_entries(
        $conf[:uri][:daily][page],
        $conf[:encoding],
        workbook.Worksheets("DailyData"))
  end
  
  # get weekly leaderboard
  pages = get_pages($conf[:uri][:weekly][1], $conf[:encoding])
  pages.each do |page|
     add_entries(
        $conf[:uri][:weekly][page],
        $conf[:encoding],
        workbook.Worksheets("WeeklyData"))
  end
  
  # save it
  workbook.Save
  workbook.Close
ensure
  excel.Quit
end
