puts 'Hello there, what\'s your name?'
Fname = gets.chomp
puts 'Cool! and your middle name,' + Fname + ' ?'
Mname = gets.chomp
puts '...and your surname?'
Sname = gets.chomp
#this is stopping it from working for some reason. must be to do with the assignment of variables
puts 'Hi ' + Fname + Mname + Sname '!' 