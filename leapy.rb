puts "Give me the two years"
 year1 = gets.chomp.to_i
 year2 = gets.chomp.to_i
 puts "This is the list of years:"

 if year1 > year2
   puts "The second year has to be bigger than the first"
 else
   while (year1 <= year2)
     if
       (((year1 % 4 == 0) and (year1 %100 !=0)) or (year1 % 400 == 0))
       puts year1.to_s

     end
     (year1 = year1.to_i + 1)
   end
puts "Finished"
end