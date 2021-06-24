from selenium import webdriver
from getpass import getpass
from selenium.common.exceptions import NoSuchElementException
import time
import xlrd
import xlsxwriter

url=input("Enter url:")
markup=input("Enter markup:")
discountp=input("Enter Discount %:")
freight=input("Enter freight price ratio:")
print("\nCreating Excel File\n")
driver= webdriver.Chrome("chromedriver.exe")
driver.get(url)
driver.maximize_window()
time.sleep(5)
name =[]
name2=[]
sku=[]
desc=[]
qty=[]
rate = []
amount = []
tradeprice=[]
subtotal=0
total=0
shipping=0
final=0
a = driver.find_element_by_xpath("//span[@id='tmp_entity_number']").text
print(a)
ouwb= xlsxwriter.Workbook(str(a)+".xlsx")
ouws1= ouwb.add_worksheet("Green Theory")
ouws= ouwb.add_worksheet("PureModern Quote")

ouws.set_margins(top=0.5, left=0.3, right=0.3,bottom=0.65)
ouws.set_header('', {'margin': 0.3})
ouws.set_footer('&C&11&"Arial,Regular"www.puremodern.com | 505 W. Riverside Ave, Suite 576, Spokane, WA 99201 | 800-563-0593', {'margin': 0.2})



bold = ouwb.add_format({'bold': True})

#for formatting headers
grey =  ouwb.add_format({
    'fg_color': '#C6C6C6',
    'text_wrap': True,
    'bold': True,
    'font_size': 13,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#999999',	
})

#for formatting headers description column
greydescription =  ouwb.add_format({
    'fg_color': '#C6C6C6',
    'text_wrap': True,
    'bold': True,
    'font_size': 13,
    'align': 'left',
    'valign': 'vcenter',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#999999',
})

#for formatting bottom totals
greybottom1=  ouwb.add_format({
    'fg_color': '#C6C6C6',
    'text_wrap': True,
    'font_size': 12,
    'align': 'right',
    'valign': 'vcenter',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#FFFFFF',
    'right': 5,
    'right_color': '#FFFFFF',
})

#for formatting bottom prices
greybottom2=  ouwb.add_format({
    'fg_color': '#C6C6C6',
    'text_wrap': True,
    'bold': True,
    'font_size': 14,
    'align': 'right',
    'valign': 'vcenter',
    'num_format': '_(\$* #,##0.00_);_(\$* (#,##0.00);_(\$* "-"??_);_(@_)',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#FFFFFF',
    'right': 5,
    'right_color': '#FFFFFF',
})
greybottom3=  ouwb.add_format({
    'fg_color': '#C6C6C6',
    'text_wrap': True,
    'font_size': 12,
    'align': 'right',
    'valign': 'vcenter',
    'num_format': '_(\$* #,##0.00_);_(\$* (#,##0.00);_(\$* "-"??_);_(@_)',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#FFFFFF'
})
bottom3=  ouwb.add_format({
    'text_wrap': True,
    'bold': False,
    'font_size': 12,
    'align': 'right',
    'valign': 'vcenter',
    'font_name': 'Arial'
})

#formatting top quote setting
merge_format = ouwb.add_format({
    'bold': 1,
    'align': 'center',
    'valign': 'bottom',
    'fg_color': 'yellow',
    'font_size': 13,
    'font_name': 'Arial',})

#formatting top Bottom Information
merge_bottom = ouwb.add_format({
    'align': 'center',
    'valign': 'bottom',
    'font_size': 13,
    'italic': True,
    'font_name': 'Arial',
    })

#for number
number=ouwb.add_format({
    'text_wrap': True,
    'font_size': 12,
    'valign': 'vcenter',
    'font_name': 'Arial',
    'align': 'left',
})

#for formatting Product Title
item=ouwb.add_format({
    'text_wrap': True,
    'bold': True,
    'font_size': 12,
    'valign': 'top',
    'font_name': 'Arial',
})

#for formatting Product Quantity, Discount
item2=ouwb.add_format({
    'text_wrap': True,
    'font_size': 12,
    'valign': 'vcenter',
    'align': 'center',
    'font_name': 'Arial',
})

#for formatting Product Prices
money2=ouwb.add_format({
    'text_wrap': True,
    'font_size': 12,
    'valign': 'vcenter',
    'align': 'center',
    'num_format': '[$$-409]#,##0.00',
    'font_name': 'Arial', 
})

subtotalfmt=ouwb.add_format({
    'text_wrap': True,
    'font_size': 12,
    'valign': 'vcenter',
    'align': 'right',
    'num_format': '[$$-409]#,##0.00',
    'font_name': 'Arial',
})


#for formatting Quote Prices
money3=ouwb.add_format({
    'text_wrap': True,
    'font_size': 11,
    'valign': 'top',
    'bold': True,
    'num_format': '[$$-409]#,##0.00',
    'font_name': 'Arial',
})

#for formatting Quote Markup, Descriptions
item3=ouwb.add_format({
    'text_wrap': True,
    'font_size': 11,
    'valign': 'top',
    'bold': True,
    'font_name': 'Arial',
})
#for aligning discount and freightratio to right
perc3=ouwb.add_format({
    'text_wrap': True,
    'font_size': 11,
    'valign': 'top',
    'bold': True,
    'font_name': 'Arial',
    'align': 'right',
})
#for formatting Quote Markup, Descriptions
perc=ouwb.add_format({
    'text_wrap': True,
    'font_size': 11,
    'valign': 'top',
    'bold': True,
    'num_format': '0.00%',
    'font_name': 'Arial',
})

#for formatting product description
item4=ouwb.add_format({
    'text_wrap': True,
    'font_size': 11,
    'align': 'left',
    'font_name': 'Arial',
    'bottom': 5,
    'bottom_color': '#999999',
})
#border format
borderf=ouwb.add_format({
    'bottom': 1,
})

#for writing input values in excel
ouws.write("I3",float(markup),item3)
ouws.write("J3",discountp+"%",perc3)
ouws.write("K3",freight+"%",perc3)
ouws1.write("A6","#",bold)
ouws1.write("B6","Item & Description",bold)
ouws1.write("C6","Qty",bold)
ouws1.write("D6","Rate",bold)
ouws1.write("E6","Amount",bold)

ouws.freeze_panes(1, 0)
ouws.set_row(0, 41)
ouws.set_column('A:A', 4.89)
ouws.set_column('B:B', 37.89)
ouws.set_column('C:C', 4.56)
ouws.set_column('D:D', 13)
ouws.set_column('E:E', 6.78)
ouws.set_column('F:F', 13)
ouws.set_column('G:G', 16.8)
ouws.set_column('H:H', 5)
ouws.set_column('I:I', 9)
ouws.set_column('J:J', 10)
ouws.set_column('K:K', 12)
ouws.set_column('L:L', 14)
ouws.set_column('M:M', 14)
ouws.set_column('N:N', 13)
ouws.set_column('O:O', 12)

ouws.write("A1", "Item",grey)
ouws.write("B1", "Description",greydescription)
ouws.write("C1", "Qty",grey)
ouws.write("D1", "Unit List Price",grey)
ouws.write("E1", "Disc. %",grey)
ouws.write("F1", "Unit Trade Price",grey)
ouws.write("G1", "Subtotal",grey)
ouws.merge_range("I1:O1","Quote Settings",merge_format)
ouws.write("I2","Mark Up",item3)
ouws.write("J2","Discount",item3)
ouws.write("K2","Freight Price Ratio",item3)
ouws.write("L2","Total Cost",item3)
ouws.write("M2","Total Price",item3)
ouws.write("N2","Margin $",item3)
ouws.write("O2","Margin %",item3)

ouws1.write("A1",a,bold)
ouws1.write("A3","Estimate Date",bold)
ouws1.write("B3","Reference#",bold)
ouws1.write("C3","Sales person",bold)

date = driver.find_element_by_xpath("//td[@id='tmp_entity_date']").text
try:
    reference = driver.find_element_by_xpath("//td[@id='tmp_ref_number']").text
except:
    reference=""
salesperson = driver.find_element_by_xpath("//td[@id='tmp_salesperson_name']").text

ouws1.write("A4",date)
ouws1.write("B4",reference)
ouws1.write("C4",salesperson)

elname = driver.find_elements_by_xpath("//span[@id='tmp_item_name']")
for n in elname:
    name.append(n.text)
print(name)
for n2 in elname:
    if n2!="Shipping.":
        name2.append(n2.text)
elsku= driver.find_elements_by_xpath("//span[@class='pcs-item-sku']")
for x in elsku:
    sku.append(x.text)
print(sku)
eldesc= driver.find_elements_by_xpath("//span[@id='tmp_item_description']")
for y in eldesc:
    desc.append(y.text)
print(desc)
elqty= driver.find_elements_by_xpath("//span[@id='tmp_item_qty']")
for z in elqty:
    qty.append(z.text)
print(qty)
elrate= driver.find_elements_by_xpath("//span[@id='tmp_item_rate']")
for az in elrate:
    temp= az.text.replace(",","")
    rate.append(temp)
print(rate)
elamount= driver.find_elements_by_xpath("//span[@id='tmp_item_amount']")
for ay in elamount:
    temp= ay.text.replace(",","")
    amount.append(temp)
print(amount)
length = len(sku)
#for First Sheet
p=6
for i in range (0,length):
    temp= sku[i]+"\n"+desc[i]
    ouws1.write(p,0,i+1)
    ouws1.write(p,1,name[i])
    ouws1.write(p+1,1,temp)
    ouws1.write(p,2,qty[i])
    ouws1.write(p,3,rate[i])
    ouws1.write(p,4,amount[i])
    p+=2

tempt=p-1
p+=2
price1 = driver.find_element_by_xpath("//td[@id='tmp_subtotal']").text
price2 = driver.find_element_by_xpath("//td[@id='tmp_total']").text
ouws1.write(p,3,"Subtotal")
ouws1.write(p,4,price1)
ouws1.write(p+1,3,"Total",bold)
ouws1.write(p+1,4,price2,bold)
totalt=p+1

k=1
m=7
n=2
p=1      #k for item and p for loop
for i in range (0,length-1):
    if(str(sku[i])!="SKU : Shipping"):
        temp= desc[i]
        ouws.write(p,0,k,number)
        if(str(sku[i])!="SKU : service-eng-drawing") and (name[i]!="Service - Shop Drawings.") and (name[i]!="Shipping."):
            ouws.write(p,1,"Modern Elite Planter - Powder Coated Aluminum",item)
        elif (name[i]=="Service - Shop Drawings."):
            ouws.write(p,1,"Shop Drawings",item)
        else:
            ouws.write(p,1,"Service - Engineer Reviewed Drawings",item)
        ouws.write(p,2,qty[i],item2)
        unitprice=float(markup)*float(rate[i])
        total+=unitprice*float(qty[i])
        tradeprice = (1-float(discountp)/100)*unitprice
        ouws.write(p,3,"=$I$3*'Green Theory'!D"+str(m),money2)
        #ouws.write(p,3,round(unitprice,2),money2)
        ouws.write(p,4,"==$J$3",item2)
        #ouws.write(p,4,str(discountp+"%"),item2)
        ouws.write(p,5,"=D"+str(n)+"-D"+str(n)+"*E"+str(n),money2)
        ouws.write(p,6,"=F"+str(n)+"*C"+str(n),subtotalfmt)
        #ouws.write(p,6,round(tradeprice*float(qty[i]),2),money2)
        subtotal+=tradeprice*float(qty[i])
        ouws.write(p+1,1,str(temp),item4)
        ouws.write(p+1,0,"",item4)
        ouws.write(p+1,2,"",item4)
        ouws.write(p+1,3,"",item4)
        ouws.write(p+1,4,"",item4)
        ouws.write(p+1,5,"",item4)
        ouws.write(p+1,6,"",item4)
        p=p+2
        k=k+1
        m=m+2
        n=n+2
subtotal=round(subtotal,2)
discount = round(total-subtotal,2)
if(str(sku[-1])=="SKU : Shipping"):
    shipping = round((float(freight)/100)*float(amount[-1]),2)
final = round(subtotal+shipping,2)
k=p
p=p+2

ouws.set_row(p, 23)
ouws.set_row(p+1, 23)
ouws.set_row(p+2, 23)
ouws.set_row(p+3, 23)
ouws.set_row(p+4, 23)

ouws.merge_range("E"+str(p+3)+":F"+str(p+3),"Subtotal",greybottom1)
ouws.merge_range("E"+str(p+2)+":G"+str(p+2),"")
#ouws.write(p+1,5,"Subtotal",greybottom1)
ouws.write(p+2,6,"=SUM(G2:G"+str(k)+")",greybottom3)
ouws.merge_range("E"+str(p+1)+":F"+str(p+1),'=TEXT(J3,"0% ")& "Total Discount"',greybottom1)
#ouws.write(p,5,discountp+"% Discount",greybottom1)
ouws.write(p,6,"=(G"+str(p+3)+"/(1-J3)-G"+str(p+3)+")*-1",greybottom3)
ouws.merge_range("E"+str(p+4)+":F"+str(p+4),"Shipping",greybottom1)
#ouws.write(p+2,5,"Shipping",greybottom1)
ouws.write(p+3,6,"='Green Theory'!E"+str(tempt)+"*'PureModern Quote'!K3",greybottom3)
ouws.merge_range("E"+str(p+5)+":F"+str(p+5),"Total USD",greybottom2)
#ouws.write(p+3,5,"Total USD",greybottom1)
ouws.write(p+4,6,"=SUM(G"+str(p+3)+":G"+str(p+4)+")",greybottom2)
#=SUM(G"+str(p+2)+":G"+str(p+2)
temptp=p+5
ouws.set_row(p+5, 17.4)
ouws.merge_range("A"+str(p+7)+":G"+str(p+7),"(Scroll down to the next page to view the terms and conditions and to accept the quote.)",merge_bottom)
if "," in price2:
    price2=price2.replace(",","")
price2=price2.replace("$","")
margin=final-float(price2)
marginp=(margin/final)*100


ouws.write("L3","='Green Theory'!E"+str(totalt+1),item3)
ouws.write("M3","=G"+str(temptp),money3)
ouws.write("N3","=M3-L3",money3)
ouws.write("O3","=N3/M3",perc)

driver.close()
while True:
    try:
        ouwb.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        # For Python 3 use input() instead of raw_input().
        decision = input("Exception caught in workbook.close(): %s\n"
                             "Please close the file if it is open in Excel.\n"
                             "Try to write file again? [Y/N]: " % e)
        if decision != 'n':
            continue

    break

print("\nExcel File "+str(a)+" Created")
