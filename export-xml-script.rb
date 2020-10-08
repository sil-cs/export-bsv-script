require 'docx'
require 'builder' # XML builder
require 'fastimage' # working with images

class String
  def is_integer?
    self.to_i.to_s == self
  end
end

if ARGV.empty?
  puts "At least one parameter required!"
  exit
end

filepath = ARGV[0]
Dir.chdir(File.dirname(filepath))
file = File.basename(filepath)

# get the docx which holds a table with all the info for the story
doc = Docx::Document.open(file)
tables = doc.tables
# try to get the directory name
dir_name = ""
if (tables[0].rows.length > 2 && tables[0].rows[1].cells.length > 3)
  dir_name = tables[0].rows[1].cells[2].to_s
end

# get all the files to find the video
all_files = Dir.entries(dir_name)
if (!all_files || all_files.length == 0)
  puts "Template folder not found. Make sure template folder is in same folder as the docx"
  exit
end

# get the total number of scenes that will be created
# this number is used in the XML header
scenes = 0
tables.each do |table| 
  table.rows.each do |row|
    page = row.cells[0].nil? ? "" : row.cells[0].to_s
    narration_text = row.cells[2].nil? ? "" : row.cells[2].to_s
    next if !page.is_integer?
    next if narration_text == ""
    scenes += 1
  end
end

# XML file builder
xml = Builder::XmlMarkup.new( :indent => 2 )
# declare the headers and the main attributes for the story project
xml.instruct! :xml, :version => "1.0", :encoding => "UTF-8" 
xml.MSPhotoStoryProject(
  :xmlns => "MSVideoStory",
  :schemaVersion => "2.0",
  :appVersion => "3.0.1115.0",
  :linkOnly => "0",
  :visualUnitCount => scenes,
  :codecVersion => "{50564D57-0000-0010-8000-00AA00389B71}",
  :sessionSeed => "85755559"
) do |x|
  # walk through each row of each table, which consists of
  # page #, narration text, scripture reference, start time, end time, translation
  prev_end_time = "0"
  tables.each do |table|

    table.rows.each do |row|
      page = row.cells[0].nil? ? "" : row.cells[0].to_s
      narration_text = row.cells[2].nil? ? "" : row.cells[2].to_s
      start_time = row.cells[4].nil? ? "" : row.cells[4].to_s
      end_time = row.cells[5].nil? ? "" : row.cells[5].to_s

      # skip this iteration if there's no narration text
      next if narration_text == ""

      # if page is a number, we have a new scene
      if page.is_integer?
        # loop through all files to find matching image and wav
        images = all_files.select do |file|
          file.downcase == page + ".jpg" || file.downcase == page + ".png"
        end
        waves = all_files.select do |file|
          file.downcase == "narration" + page + ".wav" || file.downcase == "narration" + page + ".mp3"
        end
        # check for images that matched the page/scene number
        # TODO: what do we do if there were no images or FASTIMAGE fails?
        x.VisualUnit do |vu|
          # add the timestamp
          if (start_time == "") then start_time = prev_end_time end
          vu.Timestamp :useMillis => false, :start => start_time.gsub(/[[:space:]]/, ''), :end => end_time.gsub(/[[:space:]]/, '')
          # check for an image
          if (images.length > 0)
            slide_file = images[0]
            path = File.dirname(filepath) + "/" + dir_name + "/" + slide_file
            dimensions = FastImage.size(path)
            if (dimensions)
              width = dimensions[0]
              height = dimensions[1]
              # add image to the visual unit
              vu.Image :path => slide_file, :width => width, :height => height 
            end
          end
          # check for narration
          if (waves.length > 0)
            # add narration to the visual unit
            vu.Narration :path => waves[0]
          end
        end

        prev_end_time = end_time

      end
    end # end of looping through table
  end

  # check for video file
  videos = all_files.select do |file|
    file.downcase.include? ".mp4"
  end
  if (videos.length > 0)
    x.Video :path => videos[0]
  end
end

# write the xml
File.open(File.join(dir_name, 'project.xml'), "w+") do |f|
  f.write(xml.target!)
end
