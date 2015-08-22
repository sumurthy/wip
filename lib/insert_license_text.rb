TOP_COMMENT = 'Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.'
BOTTOM_COMMENT = %{
OfficeJS Snippet Explorer, https://github.com/OfficeDev/office-js-snippet-explorer

Copyright (c) Microsoft Corporation
All rights reserved.

MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
}

NEWLINE = "\n"
JS_COMMENT_START = "/*"
JS_COMMENT_END = '*/'
HTML_COMMENT_START = "<!--"
HTML_COMMENT_END = '-->'

INPUT = 'C:/Users/suramam/git/office-js-snippet-explorer'

@jsFiles = 0 
@htmlFiles = 0

def sandwich_content(dirname) 
  Dir.foreach(dirname) do |dir|
    dirpath = dirname + '/' + dir
    if dirpath.include?('.git')
    	 puts "GIT STUFF: #{dirpath}"
    	next 
	end

    if File.directory?(dirpath) then
      if dir != '.' && dir != '..' then
        puts "DIRECTORY: #{dirpath}" ; sandwich_content(dirpath)
      end
    else
 		begin
			content = File.read(dirpath)
			# Skip if the comment is already present. 
			if content.include?('Copyright')
				next
			end
		rescue => err
			puts("READ ERROR!!!!!   #{dirpath}")
		end
		if  dirpath.end_with?(".js")
			puts "JS: #{dirpath}"
			content = JS_COMMENT_START + TOP_COMMENT + JS_COMMENT_END + NEWLINE + content + NEWLINE + JS_COMMENT_START + BOTTOM_COMMENT + JS_COMMENT_END
			File.write(dirpath, content)
			@jsFiles = @jsFiles + 1
		elsif dirpath.end_with?(".html")
			puts "HTML: #{dirpath}"
			content = HTML_COMMENT_START + TOP_COMMENT + HTML_COMMENT_END + NEWLINE + content + NEWLINE + HTML_COMMENT_START + BOTTOM_COMMENT + HTML_COMMENT_END
			File.write(dirpath, content)
			@htmlFiles = @htmlFiles + 1
		else
			puts "FILE: #{dirpath}"
		end
  		
    end
  end
end

sandwich_content(INPUT)
puts 
puts "Summary: Added license to - JS: #{@jsFiles}, HTML: #{@htmlFiles} files"

