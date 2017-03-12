
puts "Hey Sonny! It's your lovely Grandmother! How are you?"

while (response = gets.chomp) != "BYE"
    if response != response.upcase
      puts "Huh?! I CAN'T HEAR YOU!"
    end

    if (response == response.upcase)
      puts "NO! NOT SINCE " + (1930+rand(21)).to_s + "!"
    end
  end
  puts "GOOD BYE, SONNY!"