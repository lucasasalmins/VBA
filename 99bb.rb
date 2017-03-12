puts 'time to sing 99 bottles'.capitalize
bottles = 99
while bottles !=0
	puts bottles.to_s + ' bottles of beer on the wall,' + bottles.to_s + ' bottles of beer.'
	bottles = bottles -1
	puts 'Take one and pass it around,' + bottles.to_s + ' bottles of beer on the wall.'
	puts ''
end
if bottles == 0 
	puts 'No more beer, and we\'re not drunk yet'
	end
