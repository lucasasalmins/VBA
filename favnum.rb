puts 'Hello there, what\'s your favourite number?'
favnum = gets.chomp
puts 'Cool! ' + favnum.to_s + ' is a cool number'
bigger_num = favnum.to_i * 25
puts 'But ' + bigger_num.to_s + ' is WAY BIGGER and BETTER!'
