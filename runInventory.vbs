set xlApp = CreateObject("Excel.Application")

xlApp.Visible = True

set xlBook = xlApp.Workbooks.Open("\\TIER5SERVER\Tier 5 Labs\Lab Documents\SRM and Certified GMIS Inventory.xlsm")

xlBook.Save
xlBook.Close(False)
xlApp.Quit