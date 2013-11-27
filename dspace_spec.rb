require_relative './spec_helper'

describe "Dspace" do 
  it "can read from excel file "	do
    oo = Roo::Excel.new("test.xls")
    oo.default_sheet = oo.sheets.first
    oo.cell(4,'A').should == 1

    #CSV.open("result.csv","wb") do |csv|
      
    #end
  end
  it "can read file thieu from path" do 
    files = []
    File.readlines("filethieu.txt").map {|line| files << line.strip}
    files[0].should == '1_BuiDucTrong_DCL501.pdf'
  end
end