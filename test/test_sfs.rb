require './lib/odf-report'
require 'faker'

tax_rec = {
  'Accounting Profit/Loss' => rand(-20..20),
  'Franking Credits' => rand(-20..20),
  # 'Capital Gain' => rand(-20..20),
  'Non-Deductible Expenditure' => rand(-20..20),
  # 'R&D Expenditure' => rand(-20..20),
  # 'Other TOFA Income' => rand(-20..20),
  # 'Other Assessable Income' => rand(-20..20),
  # '46FA Deductions' => rand(-20..20),
  'Depreciation' => rand(-20..20),
  # 'SMB Tax Breaks' => rand(-20..20),
  # 'Capex Deductions' => rand(-20..20),
  # 'Project Pool Deductions' => rand(-20..20),
  # '40-880 Deductions' => rand(-20..20),
  # 'Other Specific Deductions' => rand(-20..20),
  # 'Exempt Income' => rand(-20..20),
  'Other Deductible Expenses' => rand(-20..20),
  'Other Non-Assessable Income' => rand(-20..20),
  # 'Losses Used' => rand(-20..20),
  # 'Losses Transferred In' => rand(-20..20),
  # 'Other TOFA Deductions' => rand(-20..20),
  'Other Adjustments' => rand(-20..20),
  'Taxable Income/Loss' => 0
}

fields = "CLIENT_NAME = British American Tobacco (Australasia Holdings) Pty Limited
END_DATE = 31 December 2013
ADDRESSED_TO_FULL = Mr Saminda Fernando
ADDRESSED_TO_POSITION = Area Tax Manager
ADDRESS_1 = 166 William Street
ADDRESS_2 = WOOLLOOMOOLOO NSW 2011
SEND_DATE = 5 September 2014
ADDRESSED_TO_FIRST_NAME = Saminda
CLIENT_INITIALS = BATAHPL
STATEMENT_OF_WORK_DATE = 6 June 2014
NPBT = 1,307,142,695
TAX_ADJUSTMENTS = 277,546,848
TAXABLE_INCOME = 1,029,595,847
TAX_AT_30_PERCENT = 308,878,754.10
LESS_FOREIGN_INCOME = 1,570,314
TOTAL_PAYG = 229,705,039
LESS_FINAL_TAX_PAYMENT = 79,860,687
REFUND_DUE = 2,257,286
PAYG_1 = 59,259,269
PAYG_2 = 51,972,922
PAYG_3 = 55,922,739
PAYG_4 = 62,550,109
CLIENT_MF_INITIALS = BATMA
CLIENT_MF_NAME = British American Tobacco Manufacturing Australia Pty Limited
CLIENT_AU_INITIALS = BATA
CLIENT_ALL_INITIALS = BAT
CLIENT_ALL_DOTS = B.A.T
DEED_DATE = 2 June 2014
CAPITAL_LOSSES = 118,128,953
REQUEST_DEED_DATE = 25 August 2014
ASSESSABLE_INCOME = 54,930,126
CAPITAL_GAIN = 74,489,854
CAPITAL_LOSSES_OFFSET = 1,249,686
NOTIONAL_CAPITAL_LOSSES = 45,569,534
DIVIDENDS = 224.6 million
TAX_RETURN_ADJUSTMENT_1 = 3,800,989
TAX_RETURN_ADJUSTMENT_2 = 2,799,000
MF_FOREIGN_SUBSIDIARIES = PNG, NZ, Fiji, Samoa and Solomon Islands"




report = ODFReport::Report.new("test/templates/temp_sfs.docm") do |r|

  fields.split("\n").each do |field|
    field.scan(/(.*) = (.*)/).each do
      r.add_field($1, $2.to_s)
    end
  end

  r.add_chart("TAX_RECONCILIATION", tax_rec)

end

report.generate("test/result/test_sfs.docm")
