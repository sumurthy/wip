require 'json'
# OpenStruct is not included by default so you have to add it.
require 'ostruct'
require_relative 'IntellisenseFromMarkdown/resource'

json_string = '{"name":"worksheet", "xmethod":[{"name":"add", "args":[{"n":10}]}]}'
json_object = JSON.parse(json_string, object_class: OpenStruct)
#puts json_object.xmethod[0].args[0].n


#json_string = '{"firstname":"jeff", "lastname":"durand", "address": { "street":"22 charlotte rd", "zipcode":"01013", "residents": 3 }}'
#puts json_object.address.street

puts "-------------------"

myResource = SpecMaker::Resource.new ({name: "worksheet", description: "Just a fancy sheet", methods:[{methodName: "Add"}, {methodName: "fancyAdd"}]})

myResourceArray = [myResource, myResource]

#puts JSON.pretty_generate myResource 

puts JSON.pretty_generate myResource