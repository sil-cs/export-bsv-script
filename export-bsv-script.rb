require 'docx'
require 'optparse'

class String
  def is_integer?
    self.to_i.to_s == self
  end
end

ARGV << "-h" if ARGV.empty?

$options = { :title => 'Title', :passage => '' }
OptionParser.new do |opts|
  opts.banner = 'Usage: export_bsv_script.rb [options] FILENAME'
  opts.on('-t', '--title TITLE', 'Title of script') do |title|
    $options[:title] = title
  end
  opts.on('-p', '--passage PASSAGE', 'Passage of script') do |title|
    $options[:title] = title
  end
  opts.on_tail("-h", "--help", "Show this message") do
    puts opts
    exit
  end
end.parse!

if ARGV.empty?
  puts opts.help
  puts "At least one parameter required!"
  exit
end

filepath = ARGV[0]
Dir.chdir(File.dirname(filepath))
file = File.basename(filepath)
title = $options[:title]
passage = $options[:passage]

# get the docx which holds a table with all the info for the story
doc = Docx::Document.open(file)
table = doc.tables[0]

# loop through table
table.rows.each do |row|
  page = row.cells[0].nil? ? "" : row.cells[0].to_s
  if page == "T"
    title = row.cells[2].nil? ? title : row.cells[2].to_s
    passage = row.cells[3].nil? ? passage : row.cells[3].to_s
  elsif page.is_integer?
    Dir.mkdir(title)  unless File.exists?(title)
    text = row.cells[2].nil? ? "" : row.cells[2].to_s
    ref = row.cells[3].nil? ? "" : row.cells[3].to_s

    File.open(File.join(title, page + ".txt"), "w+") do |f|
      f.write("#{title.strip}~#{passage.strip}~#{ref.strip}~#{text.strip}")
    end
  end
end