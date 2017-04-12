require 'docx'
require 'optparse'

class String
  def is_integer?
    self.to_i.to_s == self
  end
end

$options = { :title => 'Title' }
OptionParser.new do |opts|
  opts.banner = 'Usage: export_bsv_script.rb [options]'
  opts.on('-t', '--title TITLE', 'Title of script') do |title|
    $options[:title] = title
  end
end.parse!

filepath = ARGV[0]
Dir.chdir(File.dirname(filepath))
file = File.basename(filepath)
title = $options[:title]

doc = Docx::Document.open(file)
table = doc.tables[0]
table.rows.each do |row|
  page = row.cells[0].nil? ? "" : row.cells[0].to_s
  if page == "T"
    title = row.cells[2].nil? ? title : row.cells[2].to_s
  elsif page.is_integer?
    Dir.mkdir(title)  unless File.exists?(title)
    text = row.cells[2].nil? ? "" : row.cells[2].to_s
    ref = row.cells[3].nil? ? "" : row.cells[3].to_s

    File.open(File.join(title, page + ".txt"), "w+") do |f|
      f.write("#{title}~#{ref}~#{text}")
    end
  end
end