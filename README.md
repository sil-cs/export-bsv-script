# export-bsv-script

Export Bible Story Video Script (in docx format) to files needed by the Story Producer App.
Files will be created in a directory based on the title (command-line argument).

Usage:
```
ruby export-bsv-script.rb [-t "Title"] filename.docx
```

## Installation

### Windows Install Requirements
* Install the latest [Ruby](http://rubyinstaller.org/downloads/)
  * Check "Add Ruby executable to your PATH"
  * Check "Associate .rb and .rbw files with this Ruby installation"
* `gem install bundler`
  * With some versions of Ruby, you may get this error `Unable to download data from https://rubygems.org/ - SSL_connect returned=1...` 
  * do the following [workaround](https://gist.github.com/eyecatchup/20a494dff3094059d71d) from a cmd window (not git bash)

```
# 1. Add insecure source
> gem sources -a http://rubygems.org/
https://rubygems.org is recommended for security over http://rubygems.org/

Do you want to add this insecure source? [yn]  y
http://rubygems.org/ added to sources

# 2. Remove secure source
> gem sources -r https://rubygems.org/
https://rubygems.org/ removed from sources

# 3. Update source cache
> gem sources -u
source cache successfully updated
```
* `git clone https://github.com/chrisvire/export-bsv-script`
* `cd export-bsv-script`
* `bundle install`
