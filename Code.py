import xlrd3 as xlrd
import xlsxwriter


## Name of files we are reading information from
datafile="Invoices Info.xls"
account_activity= "GL account info.xls"

rbook = xlrd.open_workbook(datafile)
rbook_act= xlrd.open_workbook(account_activity)


## Name of file we are outputting information
output_file= "Output.xlsx"

wbook= xlsxwriter.Workbook(output_file)


## Excel Sheet Name
excel_sheet= wbook.add_worksheet("Sales Analysis")

sheet= rbook.sheet_by_index(0)
sheet_act= rbook_act.sheet_by_index(0)



def data_list(sheet):
    data= []
    row=0
    for row in range(sheet.nrows):
        col=0
        elem_data=[]
        for col in range(sheet.ncols):
            value= sheet.cell_value(row,col)
            if col == 0 and not(isinstance(value,str)):
                elem_data= elem_data + [str(round(value))]
            else:
                elem_data= elem_data + [sheet.cell_value(row,col)]
            col= col+1
        data= data + [elem_data]
        row= row + 1
    return data

lol_data_st= data_list(sheet)[1:]
lol_data_act= data_list(sheet_act)[5:]



def extract_inv_data(data,row):
    db_amt= data[row][8]
    cr_amt= data[row][7]

    if not(cr_amt == "" or db_amt == ""):
        amt= db_amt - cr_amt
    else:
        amt=''

    inv_data= data[row][4].split()
    inv=0
    for i in range(len(inv_data)):
        if inv_data[i].isnumeric:
            inv= inv_data[i]
            
    return [str(inv),amt]



def act_rows(data_act):
    data= data_act[1:]
    for i in range(len(data)-1):
        if data[i][1] == "** Closing Balance **":
            return i
        i= i+1



def inv_lst(data,count):
    inv=[]
    for i in range(count):
        inv= inv + [extract_inv_data(data,i+1)]
    return inv



def accounts_data(lol_data,acc):
    if len(lol_data) <= 1:
        return acc
    else:
        count= act_rows(lol_data)
        data= lol_data[:count+1]
        names= [data[0][0],data[0][1]]
        invoices= inv_lst(lol_data[:count+1],count)
        acc= acc + [[names] + [invoices]]
        return accounts_data(lol_data[count+4:],acc)

accounts_data_lst= accounts_data(lol_data_act,[])



def accounts(data):
    res=[]
    for i in range(len(data)):
        string= data[i][0][0] + "- " + data[i][0][1]
        res= res + [[string]]
    return res



def invs_by_acct(data):
    lst= data[1]
    res=[]
    for i in range(len(lst)):
        res= res + [data[1][i][0]]
    return res



def amts_by_acct(data):
    lst= data[1]
    res=[]
    for i in range(len(lst)):
        res= res + [data[1][i][1]]
    return res



def dicts(data):
    accounts_lst= accounts(data)
    res=[]
    num=0
    for names in accounts_lst:
        description= data[num][0][1]
        invoices= invs_by_acct(data[num])
        amts= amts_by_acct(data[num])
        name= {"description":description,"invoices":invoices,"amounts":amts}
        res= res + [name]
        num=num+1
    return res

all_dicts_data= dicts(accounts_data_lst)




## Sales Tax Data
def invoices_to_find(data):
    res=[]
    for i in range(len(data)):
        res= res + [str(data[i][0])]
    return res


## Assigning sales type and their amounts.
def sales_type_amount_assign(inv,accts_data):
    for k in range(len(accts_data)):
        inv_lst= accts_data[k]["invoices"]
        if inv in inv_lst:
            ind= inv_lst.index(inv)
            amt= accts_data["amounts"][ind]
            desc= accts_data["description"]
            return [desc,amt]



## Outputting info:
def shipping_index(accts_data):
    data_lst= accts_data
    for i in range(len(accts_data)):
        if accts_data[i]["description"] == "Sales - Shipping":
            return i
   


## WRITING
date_format = wbook.add_format({'num_format': 'mm/dd/yy'})
bold= wbook.add_format({"bold":True,"underline":True})


def output_general(sales_data):
    for r in range(len(sales_data)):
        for c in range(len(sales_data[r])):
            if c == 2:
                excel_sheet.write(r+1,c, sales_data[r][c], date_format) 
            else:
                excel_sheet.write(r+1,c,sales_data[r][c])
    


def output_sales_analysis(sales_data,accts_data):
    ind= shipping_index(accts_data)
    data= accts_data[:ind] + accts_data[ind+1:]
    inv_lst= invoices_to_find(sales_data)
    res=6

    for i in range(len(inv_lst)):
        inv= str(inv_lst[i])

        col_sale=6
        for k in range(len(data)):
            invoices= data[k]["invoices"]
            if inv in invoices:
                col_sale= col_sale + 2
                ind= invoices.index(inv)
                amt= data[k]["amounts"][ind]
                desc= data[k]["description"]
                excel_sheet.write(0,col_sale,"Sales Type",bold)
                excel_sheet.write(0,col_sale+1,"Amount",bold)
                excel_sheet.write(i+1,col_sale,desc)
                excel_sheet.write(i+1,col_sale+1,amt)

        if col_sale > res:
            res= col_sale
    
    return res



def output_shipping(sales_data,accts_data,col):
    ind= shipping_index(accts_data)
    shipping= accts_data[ind]

    inv_lst= invoices_to_find(sales_data)
    excel_sheet.write(0,col,"Shipping",bold)
    excel_sheet.write(0,col+1,"Customer Type",bold)
    for i in range(len(inv_lst)):
        inv= inv_lst[i]
        if inv in shipping["invoices"]:
            invoices= shipping["invoices"]
            amounts= shipping["amounts"]
            count= invoices.count(inv)
            amt= 0
            inc=0
            for s in range(count):
                ind= invoices.index(inv)
                amt= amt + amounts[ind]
                inc=ind
                amounts= amounts[ind+1:]
                invoices= invoices[ind+1:]
            
            excel_sheet.write(i+1,col,amt)



def output_title():
    title= data_list(sheet)[0]
    title= title 
    for i in range(len(title)):
        excel_sheet.write(0,i,title[i],bold)


def output_cust_type(sales_data,col):
    for i in range(len(sales_data)):
        if sales_data[i][4] != "CA":
            excel_sheet.write(i+1,col,"Foreign")


## Running the first set of output functions
output_general(lol_data_st)
col= output_sales_analysis(lol_data_st, all_dicts_data)
output_shipping(lol_data_st, all_dicts_data,col+2)
output_title()
output_cust_type(lol_data_st, col+3)

wbook.close()



## Reading the file we wrote in
rwbook = xlrd.open_workbook(output_file)

def data_list(sheet):
    data= []
    row=0
    for row in range(sheet.nrows):
        col=0
        elem_data=[]
        for col in range(sheet.ncols):
            elem_data= elem_data + [sheet.cell_value(row,col)]
            col= col+1
        data= data + [elem_data]
        row= row + 1
    return data

out_sheet= rwbook.sheet_by_index(0)
sheet_titles= data_list(out_sheet)[0]
complete_data= data_list(out_sheet)[1:]



## Writitng on the same file again
wbook_v2= xlsxwriter.Workbook(output_file)
bold_v2= wbook_v2.add_format({"bold":True,"underline":True})
date_format_v2 = wbook_v2.add_format({'num_format': 'mm/dd/yy'})
num_fmt = wbook_v2.add_format({'num_format': '#,##0.00'})



## Sizing Columns
def formatting(sheet):
    sheet.set_column(0,0,18.5)
    sheet.set_column(1,1,23)
    sheet.set_column(2,4,12.5)
    sheet.set_column(5,6,7.8)
    sheet.set_column(7,7,12)
    sheet.set_column(8,10,11)
    sheet.set_column(11,12,10.8)
    sheet.set_column(13,13,12.2)



num_states= int(input("How many states do you want in as seperate sheets? "))
sep_states= []


for i in range(num_states):
    inp= input("Enter State #" + str(i+1) + ": ")
    sep_states= sep_states + [inp]


def acct_names(data):
    res=[]
    for i in range(len(data)):
        res= res + [data[i][0][1]]
    return res



def totals_GL(accounts_data):
    res=[]
    for i in range(len(accounts_data)):
        sums= 0
        for s in range(len(accounts_data[i][1])):
            if not(accounts_data[i][1][s][1] == ""):
                sums= sums + accounts_data[i][1][s][1]
        res= res + [[accounts_data[i][0][1],sums]]
    return res



def ship_total_report(data):
    col= len(data[0]) - 2
    res=0
    e=0
    for i in range(len(data)):
        if data[i][col] != "":
            res= res + data[i][col]
            e=e+1
    return res



def totals_report(all_data):
    names= acct_names(accounts_data_lst)
    total_dict={}
    for i in range(len(names)):
        total_dict[names[i]] = 0

    num_cols= round((len(all_data[0]) - 10) / 2)

    for i in range(len(all_data)):
        col= 8
        for p in range(num_cols):
            if all_data[i][col] in names:
                total_dict[all_data[i][col]] = total_dict[all_data[i][col]] + all_data[i][col+1]
            else:
                p= p+1
            col= col + 2
    
    total_dict["Sales - Shipping"]= ship_total_report(all_data)

    return total_dict



def foreign_totals(all_data):
    names= acct_names(accounts_data_lst)
    total_dict={}
    for i in range(len(names)):
        total_dict[names[i]] = 0

    num_cols= round((len(all_data[0]) - 10) / 2)

    for i in range(len(all_data)):
        if all_data[i][-1] == "Foreign":
            col= 8
            for p in range(num_cols):
                if all_data[i][col] in names:
                    total_dict[all_data[i][col]] = total_dict[all_data[i][col]] + all_data[i][col+1]
                else:
                    p= p+1
                col= col + 2
        else:
            i=i+1
    return total_dict



def foreign_ship_total(all_data):
    res=0
    for i in range(len(all_data)):
        if all_data[i][-1] == "Foreign":
            if all_data[i][-2] != "":
                res= res + all_data[i][-2]
    return res                



def GL_totals_output(row,col,sales_sheet):
    sums= totals_GL(accounts_data_lst)
    total=0
    for i in range(len(sums)):
        sales_sheet.write(row,col,sums[i][0],num_fmt)
        sales_sheet.write(row,col+2,sums[i][1],num_fmt)
        total= total + sums[i][1]
        row=row+1
    sales_sheet.write(row,col+2,total,num_fmt)



def CA_totals_title_output(row,sheet):
    sheet.set_column(2,7,13)
    sheet.write(row,1,"Sales Type")
    sheet.write(row,3,"Total")
    sheet.write(row,4,"Foreign")
    sheet.write(row,5,"CA Distributor")
    sheet.write(row,6,"CA Taxable")
    sheet.write(row,7,"CA Service")



def foreign_totals_output(row,sales_sheet):
    foreign_sums= foreign_totals(complete_data)
    CA_totals_title_output(row,sales_sheet)
    row= row+1
    GL_totals_output(row,1,sales_sheet)

    foreign_total=0

    foreign_sums["Sales - Shipping"]= foreign_ship_total(complete_data)
    for key in foreign_sums:
        sales_sheet.write(row,1,key)
        sales_sheet.write(row,4,foreign_sums[key],num_fmt)
        foreign_total= foreign_total + foreign_sums[key]
        row=row+1
    
    sales_sheet.write(row,4,foreign_total,num_fmt)



def output_sheets(data,sep_states,titles):
    res=[]
    for state in sep_states:
        add_sheet= wbook_v2.add_worksheet(state)
        add_sheet.write(0,0,state + " Sales Tax Remittance")
        formatting(add_sheet)  
        for i in range(len(titles)):
            add_sheet.write(2,i,titles[i],bold_v2)
        count=0

        for i in range(len(data)):
            if data[i][4] == state:
                count=count+1
                for c in range(len(data[i])):
                    if c == 2:
                        add_sheet.write(count+2,c,data[i][c],date_format_v2)
                    else:
                        add_sheet.write(count+2,c,data[i][c])
        
        if state == "CA":
            row= count + 4
            foreign_totals_output(row,add_sheet)




def other_data(data,sep_states):
    res=[]
    for i in range(len(data)):
        if not(data[i][4] in sep_states):
            res= res + [data[i]]
    return res



def other_output(data,titles):
    excel_sheet= wbook_v2.add_worksheet("Other")
    formatting(excel_sheet)
    excel_sheet.write(0,0,"Other Data")

    for i in range(len(titles)):
        excel_sheet.write(2,i,titles[i],bold_v2)

    f_data= other_data(data,sep_states)

    for r in range(len(f_data)):
        for c in range(len(f_data[r])):
            if c == 2:
                excel_sheet.write(r+3,c, f_data[r][c], date_format_v2) 
            else:
                excel_sheet.write(r+3,c,f_data[r][c])



sales_sheet= wbook_v2.add_worksheet("Sales Analysis")

def sales_analysis_again(titles,data,excel_sheet):
    formatting(excel_sheet)
    
    for i in range(len(titles)):
        excel_sheet.write(0,i,titles[i],bold_v2)

    for r in range(len(data)):
        for c in range(len(data[r])):
            if c == 2:
                excel_sheet.write(r+1,c, data[r][c], date_format_v2) 
            else:
                excel_sheet.write(r+1,c,data[r][c])



def report_totals_output(row,sales_sheet):
    sales_sheet.write(row,2,"GL")
    sales_sheet.write(row,3,"Report")
    sales_sheet.write(row,4,"Difference")
    sums= totals_report(complete_data)
    GL_sums= totals_GL(accounts_data_lst)
    row= row+1
    total=0
    i=0
    for key in sums:
        sales_sheet.write(row,0,key,num_fmt)
        sales_sheet.write(row,3,sums[key],num_fmt)
        diff= GL_sums[i][1] - sums[key]
        sales_sheet.write(row,4,diff,num_fmt)
        total= total + sums[key]
        row=row+1
        i=i+1
    
    sales_sheet.write(row,3,total,num_fmt)


row= len(lol_data_st) + 4

sales_analysis_again(sheet_titles, complete_data, sales_sheet)
output_sheets(complete_data,sep_states,sheet_titles)

if num_states > 0:
    other_output(complete_data,sheet_titles)

GL_totals_output(row+1,0,sales_sheet)

report_totals_output(row,sales_sheet)


wbook_v2.close()