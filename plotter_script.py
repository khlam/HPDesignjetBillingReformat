import MySQLdb, sys, re, sys, os, string, openpyxl, getopt, time

stc_list = [];#global array for stc_list
plotter_accounts = []; #global 2d array for print jobs
tc_tv_ti = [0,0,0]; #global array for total cost, total valid prints, total invalid prints
accounting_summary = []; #global array for accounting summary

def read_STC(ignore_stc):
	if(os.path.exists(ignore_stc)):	
		ignore = open(ignore_stc, 'r');
		for line in ignore:
			stc_list.append(line.replace("\n","").replace("\r","").replace("'","").replace("[","").replace("]","").lower());
	else:
		print "\nIgnore list '" + str(ignore_stc) + "'not found. Not ignoring anyone."
	

def check_STC(username, line_num, index_num, onquit):
	if username == "None" or username.replace("\r","").replace("'","").lower() in stc_list:
		if onquit == 1:
			print "\nInvalid username " + username + ", line " + str(line_num) + ", Index: " + str(index_num) + "\nScript end.";
			exit(1);
		print "\nInvalid username " + username + ", line " + str(line_num) + ", Index: " + str(index_num) + "\nEnter the correct username to continue or 'Q' to quit without saving. ";
		username = raw_input("\t>");
		if username == 'Q':
			print "Script End."
			exit(1);
		return username;
	else:
		return username;

def handle_index(index):
	index = index.replace(" ","");
	index_length = len(index);
	if (index_length == 6 or index_length == 11):
		return index;
	if (index_length == 10):
		return index[:6] + "-" + index[6:]
	if (index_length != 6 or index_length != 11):
		print "Index: " + index + " is entered incorrectly. Please manually fix."


def make_spreadsheet(spreadsheet, final_spreadsheet, onquit):
	try:
		wb = openpyxl.load_workbook(spreadsheet);
		print "\nSuccessfully opened " + spreadsheet + ". Script Starting";
	except:
		print "\nCannot load " + spreadsheet + ". Please check target file name.";
		sys.exit(2);
	ws = wb.worksheets[0];
	for row in ws.iter_rows(): #iterate through every row
		for cell in row: #iterate through every cell, wastes time but whatever
			if (cell.value == "Paper" or cell.value == "Total"): #if paper column then 
				ws.row_dimensions[cell.row].hidden = True; #hide it

				if(ws.cell(row = cell.row - 1, column = 4).value == 'printed'): #if we printed the column, then 
					ws['B'+ str(cell.row - 1)] = ws.cell(row = cell.row, column = 2).value; #move the cost up 1 cell
					if(check_STC != 0): 
						ws['J'+ str(cell.row - 1)] = check_STC(str(ws.cell(row = cell.row-1, column = 10).value), cell.row-1, ws.cell(row = cell.row-1, column = 12).value, onquit);
					try: 
						tc_tv_ti[0] = tc_tv_ti[0] + float(ws.cell(row = cell.row, column = 2).value); #add total cost
						plotter_accounts.append([]); #add another array to make plotter_accounts a 2d array
						plotter_accounts[tc_tv_ti[1]].append('{0:.2f}'.format(ws.cell(row = cell.row, column = 2).value)); #0, Document cost
						ws.cell(row = cell.row, column = 2).value = 0; #set value to zero
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 3).value)); #1, Document name
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 6).value)); #2, Paper type
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 7).value)); #3, Paper sq. ft
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 10).value)); #4, Username
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 11).value)); #5, Printing Time
						plotter_accounts[tc_tv_ti[1]].append(handle_index(str(ws.cell(row = cell.row-1, column = 12).value))); #6, Billing Index
						plotter_accounts[tc_tv_ti[1]].append(str(ws.cell(row = cell.row-1, column = 13).value)); #7, Paper Quality
						tc_tv_ti[1] = tc_tv_ti[1] + 1;
					except ValueError:
						print "Row " + str(cell.row) + " contains an invalid value. Not included in total cost."
						tc_tv_ti[2] = tc_tv_ti[2] + 1;
					ws.row_dimensions[cell.row - 1].hidden = False; #unhide the row
			else:
				ws.row_dimensions[cell.row].hidden = True; #hide everything else

	ws.row_dimensions[1].hidden = False; #except the first row with all the column labels
	ws['A' + str(ws.max_row + 1)] = 'Total';
	ws['B' + str(ws.max_row)] = '$'+str(tc_tv_ti[0]);
	wb.save(filename = final_spreadsheet) #save the document

def make_billing_detail(plotter_billing_detail):
	if(os.path.exists(plotter_billing_detail)):
		os.remove(plotter_billing_detail);
	billing_detail = open(plotter_billing_detail, 'w+');
	billing_detail.write("{\\rtf1\\ansi\deff0\n");
	billing_detail.write("\landscape\n\paperw15840\paperh12240\margl450\margr450\margt720\margb720\\tx720\\tx1440\\tx2880\\tx5760\n\\trowd\\trgaph144\n\\fs24");
	billing_detail.write("\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 User Plotter Billing\par}\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 "+time.strftime("%B %Y")+"\par}")
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx1000\n");#cost value
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx5000\n"); #document
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx6500\n"); #paper type
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx7500\n"); #paper used
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx11500\n"); #username
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx12800\n"); #printing time
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx13800\n"); #index
	billing_detail.write("\clbrdrt\\brdrs\clbrdrl\\brdrs\clbrdrb\\brdrs\clbrdrr\\brdrs\n\cellx15000\n"); #print quality
	billing_detail.write("\nCost Value\intbl\cell\nDocument\intbl\cell\nPaper Type\intbl\cell\nPaper Used\intbl\cell\nUsername\intbl\cell\nPrinting Time\intbl\cell\nIndex\intbl\cell\nPrint Quality\intbl\cell\n\\row");
	for i, item in enumerate(plotter_accounts):
		billing_detail.write("$"+str(plotter_accounts[i][0])+"\intbl\cell\n"+str(plotter_accounts[i][1])+"\intbl\cell\n"+str(plotter_accounts[i][2])+"\intbl\cell\n"+str(plotter_accounts[i][3])+"\intbl\cell\n"+str(plotter_accounts[i][4])+"\intbl\cell\n"+str(plotter_accounts[i][5])+"\intbl\cell\n"+str(plotter_accounts[i][6])+"\intbl\cell\n"+str(plotter_accounts[i][7])+"\intbl\cell\\row");
	billing_detail.write("\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 Total: $"+str(tc_tv_ti[0])+"\par}}}");
	billing_detail.close();

def sort_for_account_summary(plotter_accounts):
	previous_user = [];
	previous_index = [];
	for i, item in enumerate(plotter_accounts):
		index = str(plotter_accounts[i][6])
		user = str(plotter_accounts[i][4]) 
		if (user not in previous_user or index not in previous_index):
			accounting_summary.append([]);	
			k = len(accounting_summary)-1;
			accounting_summary[k].append(user); #0 - user
			accounting_summary[k].append(index); #1 - index number
			accounting_summary[k].append(float(0.00)); #2 - total cost
			for x, item in enumerate(plotter_accounts[0:]):
				total = 0;
				if(user == plotter_accounts[x][4] and index == plotter_accounts[x][6]):
					accounting_summary[k][2] = '{0:.2f}'.format(float(accounting_summary[k][2]) + float(plotter_accounts[x][0]));
			previous_index.append(index);
			previous_user.append(user);
	del(previous_user[:])
	del(previous_index[:])
		#	print "u: "+str(accounting_summary[k][0])+ " | " +str(accounting_summary[k][1])+ " | " +str(accounting_summary[k][2])

def make_account_summary(final_account_summary):
	if(os.path.exists(final_account_summary)):
		os.remove(final_account_summary);
	account_summary_doc = open(final_account_summary, 'w+');
	account_summary_doc.write("{\\rtf1\\ansi\deff0 {\\fonttbl}\n");
	account_summary_doc.write("\paperw11907\paperh16838\margt1000\margl1200\margb1000\margr1200\sectd\\titlepg{\headerf{\pard\plain \s0\ql\sb60\sa60\\f0\\fs24 "+time.strftime("%m/%d/%Y")+"\par}}{\header}\n")
	account_summary_doc.write("\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 Plotter Accounting\par}\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 "+time.strftime("%B %Y")+"\par}\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 Please JV the following indexes to FOR078-FPLT\par}\n")
	
	account_summary_doc.write("\n{\\trowd\\trgaph108\\trql\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx2160\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx4320\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx6480\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24 Username\cell\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24 Index\cell\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24 Total\cell\\row}\n");
	for i, item in enumerate(accounting_summary):
		account_summary_doc.write("\n{\\trowd\\trgaph108\\trql\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx2160\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx4320\clbrdrt\\brdrs\\brdrw20\clbrdrl\\brdrs\\brdrw20\clbrdrb\\brdrs\\brdrw20\clbrdrr\\brdrs\\brdrw20\cellx6480\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24"
			+str(accounting_summary[i][0])+"\cell\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24 " #username
			+str(accounting_summary[i][1])+"\cell\pard\plain\intbl \s2\ql\sb240\sa60\\f0\\fs24 $" #index
			+str(accounting_summary[i][2])+"\cell\\row}\n"); # cost
	account_summary_doc.write("\n{\pard\plain \s2\ql\sb240\sa60\\f0\\fs24 Total: $"+str(tc_tv_ti[0])+"\par}}}");
	account_summary_doc.close();

def main(argv):
	ignore_stc = 'ignore.txt';
	spreadsheet = 'test.xlsx';
	final_spreadsheet = time.strftime("%Y_%d.xlsx")
	final_account_summary = time.strftime("%Y_%d Plotter Accounting Summary.rtf")
	plotter_billing_detail = time.strftime("%Y_%d Plotter Billing Detail.rtf")
	onquit = 0;
	try: #start of command line argument check type --help for options
			opts, args = getopt.getopt(argv,"s:t:a:i:q");
	except getopt.GetoptError:
			print "\n\tDefault settings will be used for all unspecified inputs.\n\t-s\ttarget .xlsx spreadsheet\n\t-t\tname of finished .xlsx spreadsheet\n\t-a\tname of accounting summary.rtf\n\t-b\tname of plotter_billing_detail.rtf\n\t-i\tlist of staff to ignore\n\t-q\tscript will quit if it finds an STC username"
			sys.exit(2);
	for opt, arg in opts:
		if opt in ("-s"):
			spreadsheet = arg; #set target spreadsheet name
		if opt in ("-t"):
			final_spreadsheet = arg; #set final spreadsheet name
		if opt in ("-a"):
			final_account_summary = arg;
		if opt in ("-b"):
			plotter_billing_detail = arg;
		if opt in ("-i"):
			ignore_stc = arg; #set ignore list
		if opt in ("-q"):
			onquit = 1;
			print "\nONQUIT is on. Script will quit if STC user is found."

	read_STC(ignore_stc);
	make_spreadsheet(spreadsheet, final_spreadsheet, onquit);
	make_billing_detail(plotter_billing_detail);
	sort_for_account_summary(plotter_accounts);
	make_account_summary(final_account_summary);

	print "Script Finished."
	print "\n\tEdited spreadsheet saved as " + final_spreadsheet
	print "\t  Valid Prints: " + str(tc_tv_ti[1]);
	print "\tInvalid Prints: " + str(tc_tv_ti[2]);
if __name__ == "__main__": #call to main to check for command line prompts
	doc3 = main(sys.argv[1:])
	del(stc_list[:])
	del(plotter_accounts[:])
	del(tc_tv_ti[:])
	del(accounting_summary[:])
	#made by Kin-Ho Lam for COF helpdesk 5/5/16
	
	
