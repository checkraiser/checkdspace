require 'roo'  
require 'csv'

def check?(a, b)
  return false unless b
  bs = b.split(";").to_a
  bs.each do |k|
    return true if a.include?(k)
  end
  return false
end
files = []
File.readlines("filethieu.txt").map {|line| files << line.strip.upcase}

oo = Roo::Excel.new("test.xls")
temps = []
oo.sheets.each_with_index do |sheet, index|
  oo.default_sheet =  oo.sheets[index] #oo.sheets.first  
  fr =  oo.first_row
  lr = oo.last_row  
  col = (index == 0) ? 'J' : 'L'
  (fr..lr).each do |t|
    te = oo.cell(t, col).strip.upcase if oo.cell(t, col)
    if check?(files,te)  
      temp = []      
      'A'.upto(col) do |c|
        temp << oo.cell(t, c) 
      end      
      temps << temp  
    end
  end  
end
#puts temps.count
CSV.open("result.csv","wb:UTF-8") do |csv|
  temps.each do |t|
    csv << t
  end
end
  #col = (index == 0 ? 'J' : 'L')
  #oo.first_row.upto(oo.last_row) do |row_number|        
   # CSV.open("result.csv","wb:UTF-8") do |csv|
   #   if files.include?(oo.cell(row_number, 'J')) then 
    #    temp = []
     #   'A'.upto('J') do |c|
      #     temp << oo.cell(row_number, c)
      #  end
      #  csv << temp
     # end
   # end
 # end
#end
