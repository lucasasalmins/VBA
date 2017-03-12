puts 'what is your full name?'
Fname = gets.chomp
puts 'There are ' + Fname.length.to_s + ' characters in your first name'
puts 'now your surname'
Sname = gets.chomp
Totalnamecount = Fname.length.to_i + Sname.length.to_i 
puts 'there are total ' + Totalnamecount.to_s + ' characters in your name'
