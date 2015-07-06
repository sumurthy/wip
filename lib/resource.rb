require 'json'

module SpecMaker

class Resource
	attr_accessor :name, :description, :jMethods
	Method = Struct.new(:name, :params)	


	def initialize(opts = {})

    	@name = opts[:name]
    	@description = opts[:description]
    	@jMethods = []
    	opts[:methods].each do |aMethod|
    		@jMethods.push << Method.new(aMethod[:methodName], nil).to_h
    	end
    end

 # 	def to_json(*a)
 #    	#{ 'data' => @name, @description, @jMethods] }.to_json(*a)
 #    	JSON.pretty_generate (
	# 		{
	#     		name: @name,
	#     		description: @description,
	#     		methods: @jMethods.to_h
 #    		})
	# end

 	def to_json(*a)
    	#{ 'data' => @name, @description, @jMethods] }.to_json(*a)
    		{
	    		name: @name,
	    		description: @description,
	    		methods: @jMethods
    		}.to_json(*a)
	end	

	
end

end

#    	@description = opts[:description]
#    	@description = opts[:description]
#    	@methods = opts[:methods]
