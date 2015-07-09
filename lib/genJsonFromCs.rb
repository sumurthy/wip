###
# This program reads the .CS metadata file and generates the JSON representaiton as separate files for each resources. 
# Location: https://github.com/sumurthy/wip
####

module SpecMaker

require 'logger'
require 'json'
#require_relative 'resource'
LOG_FILE = '../../logs/genJsonFromCS_log.txt'

begin
	File.delete(LOG_FILE)
rescue => err
	#Ignore this error
end

@logger = Logger.new(LOG_FILE)
@logger.level = Logger::DEBUG

#EXCELAPI_FILE_SOURCE = '../../data/ExcelApi.cs'
EXCELAPI_FILE_SOURCE = '../../data/ExcelApi_june.cs'
ENUMS = 'jsonFiles/enums/enums.json'
LOADMETHOD = 'jsonFiles/enums/loadMethod.json'
JSONOUTPUT_FOLDER = 'jsonFiles/'


#JSON_OUT = 'c:/ruby/excel_json.txt'

@csarray = []
@csarray_pure = []
@csarray_out = []
@current_object = ''
@resource = {}

# Master json containers
@json_hash = {}
@json_array = []
@json_object = {}
@json_object[:name] = ''
@json_object[:description] = ''
@json_object[:isCollection] = false
@json_object[:info] = {}
@json_object[:info][:version] = '1.0'
@json_object[:info][:title] = 'Office JavaScript Add-in API'
@json_object[:info][:description] = 'Office JavaScript Add-in API for June fork'
@json_object[:properties] = []
@json_object[:methods] = []

# Sub json containers
# Method = Struct.new(:name, :returnType, :description, :parameters, :syntax, :vbaInfo, :signature)
# Property = Struct.new(:name, :dataType, :description, :isReadOnly, :enumNameJs, :isCollection, :vbaInfo, :possibleValues, :isRelationship)
Method = Struct.new(:name, :returnType, :description, :parameters, :syntax, :signature, :notes)
Property = Struct.new(:name, :dataType, :description, :isReadOnly, :enumNameJs, :isCollection, :possibleValues, :isRelationship, :notes)
ParamStr = Struct.new(:name, :dataType, :description, :isRequired, :enumNameJs, :possibleValues, :notes)

SIMPLETYPES = %w[int string object object[][] object[] double bool number void]

def csarray_write (line=nil) 
	@csarray_out.push line
end

### 
# Load up all the known existing enums. Remove leading Excel.
##
@enumHash = {}
tempEnumHash = JSON.parse File.read(ENUMS)
@enumHash = Hash[tempEnumHash.map {|k, v| [k.gsub('Excel.',''), v] }]

@loadMethodHash = {}
@loadMethodHash = JSON.parse(File.read(LOADMETHOD), {:symbolize_names => true})

### 
# Read the file & create a transit file by removing existing comments from the .CS File.
##
@csarray_pure = File.readlines(EXCELAPI_FILE_SOURCE) 
handle_getItem = ''

@csarray_pure.each do |line|
	if line.include?('this[')
		handle_getItem = line
		handle_getItem = handle_getItem[0,handle_getItem.index('{')]
		handle_getItem = handle_getItem.gsub('this[','getItem(').gsub(']',');')
		line = handle_getItem + "\n"
	end
	@csarray.push line	
end

def self.uncapitalize (str="")
	if str.length > 0
		str[0, 1].downcase + str[1..-1]
	else
		str
	end
end


### 
# Forward Pass: Write to the output array
##

# #remove from here
# @json_object[:name] = ''
# @json_object[:description] = ''
# @json_object[:properties] = []
# @json_object[:methods] = []


# #end remove

@json_objects = nil
in_region = false
member_summary = ''
member_ahead = false
readOnly = false
property_array = []
method_array = []
parm_array = []
parm_array_metadata = []
enumName = ''
parm_hash_array = []

@csarray.each_with_index do |line, i|

	## For new object, load its resource and fill the description
	if line.include?('public interface')
		# Get the third word
		@json_object[:name] = line.split.first(3).join(' ').split.last(1).join(' ').gsub(':','')
		if @json_object[:name].include?('Collection')
			@json_object[:isCollection] = true
		else
			@json_object[:isCollection] = false
		end
		in_region = true
		@member_summary = ''
	end
	
	if line.strip.start_with?('[ClientCallableComMember', '[ClientCallableOperation') 
		member_ahead = true
	end
	
	if in_region && line.start_with?("\t}")		

		# If this is a collection, add the 'items' property as that is not listed in the .CS file for some reason! 
		if 	@json_object[:isCollection] == true
			prop_name = 'items'
			isRel = false

			objectMadeOf = uncapitalize (@json_object[:name][0, @json_object[:name].index('Collection')])
			makeDesc = "A collection of " + objectMadeOf + " objects."
			readOnly = true
			enumName = nil
			isItCollection = true
			itemReturnType = @json_object[:name][0,@json_object[:name].index('Collection')] + '[]'

			property = Property.new(prop_name, itemReturnType, makeDesc, readOnly, enumName, isItCollection, nil, isRel, nil)	
			property_array.push property.to_h
			property = nil		
		end		


		# Write the buffer
		if property_array.length == 0
			@json_object[:properties] = nil
		else
			@json_object[:properties] = property_array
		end

		# Add the .load method to the method array. 		
		if property_array.length == 0
			puts @json_object[:name]
			if method_array.length == 0
				@json_object[:methods] = nil
			else
				@json_object[:methods] = method_array				
			end
		else
			method_array.push @loadMethodHash		
			@json_object[:methods] = method_array				

		end	

		File.open("#{JSONOUTPUT_FOLDER}#{(@json_object[:name]).downcase}.json", "w") do |f|
			f.write(JSON.pretty_generate @json_object)
		end

		# Reset the variables.
		in_region = false
		parm_hash_array = []
		parm_array = []
		property_array = []
		method_array = []
	end

	# Load up either the object or member summary.
	if line.include?('/// <summary>') && 
		if !in_region
			@json_object[:description] = @csarray[i+1].delete!('///').strip
		else
			member_summary = @csarray[i+1].delete!('///').strip

			if member_summary.index('See Excel.') != nil
				enumName = member_summary[member_summary.index('See Excel.')..-1].split[1]
				member_summary = member_summary[0,member_summary.index('See Excel.')-1]
				enumName = enumName.chomp('.')
			elsif member_summary.index('Refer to Excel.') != nil
				enumName = member_summary[member_summary.index('Refer to Excel.')..-1].split[2]			
				member_summary = member_summary[0,member_summary.index('Refer to Excel.')-1]				
				enumName = enumName.chomp('.')				
			else
				enumName = nil
			end

			# if member_summary.include?('Read-only')
			# 	readOnly = true
			# 	member_summary = member_summary.gsub('. Read-only.', '.')
			# 	member_summary = member_summary.gsub(' Read-only.', '.')
			# else
			# 	readOnly = false
			# end
		end
	end

	# if method has params, then load them and also mark the enumName.
	if line.include?('/// <param name=')
		param_summary = line.split('>')[1].gsub('</param', '')

		if param_summary.index('See Excel.') != nil
			enumName = param_summary[param_summary.index('See Excel.')..-1].split[1]
			param_summary = param_summary[0,param_summary.index('See Excel.')-1]
			enumName = enumName.chomp('.')
		elsif param_summary.index('Refer to Excel.') != nil
			enumName = param_summary[param_summary.index('Refer to Excel.')..-1].split[2]			
			param_summary = param_summary[0,param_summary.index('Refer to Excel.')-1]				
			enumName = enumName.chomp('.')
		else
			enumName = nil
		end	


		param_name = line.split('"')[1]
		parameter = ParamStr.new(param_name, nil, param_summary, nil, enumName, nil, nil)	
		parm_array.push parameter
	end

	# Presence of { would indicate that it is a property or a relation	
	if member_ahead && !line.include?('_') && line.include?('{')  

		prop_name = line.split[1]
		prop_name = uncapitalize prop_name
		member_ahead = false		

		if line.include?('Collection')
			isItCollection = true
		else
			isItCollection = false
		end
		if line.include?('set;')
		 	readOnly = false
		 	member_summary = member_summary.gsub('. Read-only.', '.')
		 	member_summary = member_summary.gsub(' Read-only.', '.')
		else
		 	readOnly = true
		 	member_summary = member_summary.chomp(' Read-only.')
		 	member_summary = member_summary.chomp('Read-only.')
		end

		proDataType = line.split[0].gsub('?','')

		if (proDataType != 'object[]') && (proDataType != 'object[][]')
			proDataType = proDataType.chomp('[][]')
			proDataType = proDataType.chomp('[]')		
		end

		# If the return type is primitive or one of the enums, then it's a property. else a relation.
		if (SIMPLETYPES.include? proDataType) || (@enumHash.has_key? proDataType)
			isRel = false
		else
			isRel = true
		end
		if enumName != nil
			proDataType = 'string'
		end

		if @enumHash.has_key? proDataType
			enumName = 'Excel.' + proDataType
			proDataType = 'string'
		end

		property = Property.new(prop_name, proDataType, member_summary, readOnly, enumName, isItCollection, nil, isRel, nil)	
		property_array.push property.to_h
		property = nil
	end

	# If member is a method and has param, capture its optional param and data type.
	if member_ahead && line.include?(');')  && !line.include?('();') && !line.include?('_') 

		# Capture the first part of the parameter definition inside method definition to see if it has readonly flag and also note down its data type.
		parm_array_metadata = line[line.index('('), line.index(');')].split(',')
		parm_array_metadata.map! {|n| n.split[0]}


		parm_array_metadata.each_with_index do |metadata, j|
			if metadata.include?('Opitional') || metadata.include?('Optional') 
				parm_array[j][:isRequired] = false
			else
				parm_array[j][:isRequired] = true
			end
		end
	
		# Now we can thrash the metadata array to figure out the actual data type
		parm_array_metadata.map! {|n| n.gsub('?','')}
		parm_array_metadata.map! {|n| n.gsub(']',' ')}
		parm_array_metadata.map! {|n| n.gsub(';',' ')}

		parm_array_metadata.each_with_index do |dataType, j|
			suffix = ''
			if dataType.include?('Array')  
				if dataType.include?('Array<Array')
					suffix = '[][]'
				else
					suffix = '[]'
				end
			end
			parm_array[j][:dataType] = dataType.split[-1].gsub('(','') + suffix
			if parm_array[j][:dataType] == 'int'	
				parm_array[j][:dataType] = 'number'
			end

			# If the enum still slips through to the data type, then overwrite and set the enum correctly. 
			if @enumHash.has_key? parm_array[j][:dataType] 
				parm_array[j][:enumNameJs] = 'Excel.' + parm_array[j][:dataType] 
				parm_array[j][:dataType]  = 'string'
			end

			# Enum data type should be documented as strings. 
			if enumName != nil
				parm_array[j][:dataType] = 'string'
			end
			parm_array[j][:dataType] = parm_array[j][:dataType].gsub('?', '')
		end

		parm_array.each do |parmStruct| 
			parm_hash_array.push parmStruct.to_h 
		end
	end

	# If it's a method, dump its informaiton 
	if member_ahead && !line.include?('_') && line.include?(');')  
		#Method = Struct.new(:name, :returnType, :description, :parameters)
		#@json_object[:methods] = []
		member_ahead = false
		temp = line.split[1]
		mthd_name = temp[0,temp.index('(')]
		mthd_name = uncapitalize mthd_name	
		if parm_hash_array.length == 0
			parm_hash_array = nil
			signature = mthd_name + '()'
			syntax = "#{uncapitalize @json_object[:name]}Object.#{mthd_name}();"
		else			
			signature = mthd_name + '(' 
			syntax = "#{uncapitalize @json_object[:name]}Object.#{mthd_name}("

			parm_hash_array.each_with_index do |parmhash, k|
				signature = signature + parmhash[:name] + ': ' + parmhash[:dataType] 
				syntax = syntax + parmhash[:name] 

				if k < (parm_hash_array.length - 1)
					signature = signature + ', '
					syntax = syntax + ', '
				end
			end
			signature = signature + ')'
			signature = signature.gsub('?','')
			syntax = syntax + ');'
			syntax = syntax.gsub('?','')
		end
		line.split[0] = line.split[0].gsub('?','')
		method = Method.new(mthd_name, line.split[0], member_summary, parm_hash_array,syntax, signature, nil)
		method_array.push method.to_h
		method = nil
		parm_array = []
		parm_hash_array = []
	end
end

	puts "*** Run Completed ***"

#end module
end